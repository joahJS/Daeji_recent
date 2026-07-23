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
using System.Data.SqlClient;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * 
 * --------------------HISTORY------------------
 * 1. 2021-02-07 (현업요청)
 *    거래처 초성검색 추가
 *    
 * 1. 수정일자: 2022-10-13
 *    수정자  : 정은영
 *    ID      : #0001
 *    수정내용: 지불관리 적요변경
 *    
 */
namespace AccAdm
{
    public partial class AC16001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC16001F01()
        {
            InitializeComponent();
        }

        public AC16001F00 P_AC16001F00;
        public DataRow DR_INFO;
        public string AddAndMdi { get; set; }

        private void AC16001F01_Load(object sender, EventArgs e)
        {
            SetInit();

            if(DR_INFO != null)
            {
                TxtSlipNo.EditValue = DR_INFO["SLIPNO"];
                DateEditTDate.EditValue = DR_INFO["TDATE"];
                DateEditTDate.ReadOnly = true;
                LkupJGubn.EditValue = DR_INFO["JGUBN"];
                LkupBilGu.EditValue = DR_INFO["BILGU"];
                BtnEditCvNam.EditValue = DR_INFO["DEALER_NM"];
                TxtCvCod.EditValue = DR_INFO["CVCOD"];
                BtnEditActNo.EditValue = DR_INFO["ACNT"];
                TxtAcntNo.EditValue = DR_INFO["ACTNO"];
                TxtTrSum.EditValue = DR_INFO["TRSUM"];
                TxtRk.EditValue = DR_INFO["RK"];
                TxtAText.EditValue = DR_INFO["ATEXT"];//#0001
            }
        }

        private void SetInit()
        {
            DateEditTDate.EditValue = DateTime.Today;
            TxtTrSum.EditValue = 0;
            TxtAText.EditValue = "외상대 지불";//#0001

            LkupJGubn.Properties.ValueMember = "CD";
            LkupJGubn.Properties.DisplayMember = "NM";
            LkupJGubn.Properties.DataSource = GetLookUpData("2", "Y", "");
            LkupJGubn.EditValue = "1";
            LkupBilGu.Properties.ValueMember = "CD";
            LkupBilGu.Properties.DisplayMember = "NM";
            LkupBilGu.Properties.DataSource = GetLookUpData("1", "Y", "");
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sSlipNo = TxtSlipNo.EditValue?.ToString().Trim();
            string sTDate = DateEditTDate.EditValue?.ToString().Substring(0, 10);
            string sJGbun = LkupJGubn.EditValue?.ToString();
            string sBilGu = LkupBilGu.EditValue?.ToString();
            string sCvCod = TxtCvCod.EditValue?.ToString();
            string sActNo = TxtAcntNo.EditValue?.ToString();
            double dTrSum = string.IsNullOrEmpty(TxtTrSum.EditValue?.ToString()) ? 0 : Convert.ToDouble(TxtTrSum.EditValue);
            string sRk = TxtRk.EditValue?.ToString().Trim();

            string sAcCod1 = string.Empty;
            string sAcCod2 = string.Empty;

            string sAText = string.Empty;

            if (AddAndMdi.Equals("ADD"))
            {
                sAText = BtnEditCvNam.EditValue?.ToString() + "," + TxtAText.EditValue?.ToString().Trim(); //적요 (요청)금액제거 2022-12-09
            }
            else
            {
                sAText = TxtAText.EditValue?.ToString().Trim();
            }

            if (string.IsNullOrEmpty(sTDate))
            {
                XtraMessageBox.Show("발행일자를 입력하세요.");
                DateEditTDate.SelectAll();
                DateEditTDate.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sJGbun))
            {
                XtraMessageBox.Show("결제대상을 선택하세요.");
                LkupJGubn.SelectAll();
                LkupJGubn.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sBilGu))
            {
                XtraMessageBox.Show("지급구분을 선택하세요.");
                LkupBilGu.SelectAll();
                LkupBilGu.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sCvCod))
            {
                XtraMessageBox.Show("거래처를 입력하세요.");
                BtnEditCvNam.SelectAll();
                TxtCvCod.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sActNo))
            {
                XtraMessageBox.Show("계좌를 입력하세요.");
                BtnEditActNo.SelectAll();
                TxtAcntNo.Focus();
                return;
            }
            else if(dTrSum <= 0)
            {
                XtraMessageBox.Show("지불금액을 입력하세요.");
                TxtTrSum.SelectAll();
                TxtTrSum.Focus();
                return;
            }
            //else if (string.IsNullOrEmpty(sAText))
            //{
            //    XtraMessageBox.Show("적요를 입력하세요.");
            //    TxtAText.SelectAll();
            //    TxtAText.Focus();
            //    return;
            //}

            if (sBilGu.Equals("0001"))
            {
                sAcCod2 = "0103"; //송금의 경우 계정코드 0103 HARDCODING -> 계정코드 변경 시 해당 코드도 변경하여야함
            }
            else if (sBilGu.Equals("0002"))
            {
                sAcCod2 = "0101"; //현금의 경우 계정코드 0101 HARDCODING;
            }

            /*
             * COM_BASE_CD 테이블 -> CD_GB : AC16001_01
             * COM_CD -> 1 : 외상매입금
             * COM_CD -> 2 : 미지급금
             * COM_CD -> 3 : 외상매출금
             * COM_CD -> 4 : 미수금
             */
            if (sJGbun.Equals("1"))
            {
                sAcCod1 = "0251";
            }
            else if (sJGbun.Equals("2"))
            {
                sAcCod1 = "0253";
            }

            /*
             * 전표번호가 존재하는 경우 전표테이블 조회하여 승인여부체크
             */
            if (!string.IsNullOrEmpty(sSlipNo))
            {
                StringBuilder strSql1 = new StringBuilder();

                strSql1.Clear();
                strSql1.AppendLine(" SELECT COUNT(CASE WHEN APVYN = 'Y' THEN 1 END) AS CNT  ");
                strSql1.AppendLine("   FROM ACTRAN A ");
                strSql1.AppendLine("  WHERE A.AAUTO = 'D01' ");
                strSql1.AppendLine("    AND A.REF1 = '" + sSlipNo + "' ");
                strSql1.AppendLine("  GROUP BY A.TDATE, A.ATGUB, A.SEQNO ");

                DataTable dtCnt = DBConn.GetDataTable(DBConn.dbCon, strSql1.ToString());
                int iCnt = 0;
                if (dtCnt.Rows.Count > 0)
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
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                //지불번호가 존재하지 않는 경우
                //전표 INSERT
                if (string.IsNullOrEmpty(sSlipNo))
                {
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
                    sSlipNo = cmd.ExecuteScalar()?.ToString();

                    //dicParams.Add("@SLIPNO", sSlipNo);
                    //dicParams.Add("@TDATE", sTDate);
                    //dicParams.Add("@JGUBN", sJGbun);
                    //dicParams.Add("@BILGU", sBilGu);
                    //dicParams.Add("@CVCOD", sCvCod);
                    //dicParams.Add("@ACTNO", sActNo);
                    //dicParams.Add("@TRSUM", dTrSum.ToString());
                    //dicParams.Add("@RK", sRk);
                    //dicParams.Add("@CUSER", FmMainToolBar2.UserID);

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
                    strSql.AppendLine("     VALUES( '"+ sSlipNo + "' ");
                    strSql.AppendLine("           , '"+ sTDate + "' ");
                    strSql.AppendLine("           , '"+ sJGbun + "' ");
                    strSql.AppendLine("           , '"+ sBilGu + "' ");
                    strSql.AppendLine("           , "+ sCvCod + " ");
                    strSql.AppendLine("           , "+ sActNo + " ");
                    strSql.AppendLine("           , "+ dTrSum + " ");
                    strSql.AppendLine("           , '"+ sRk + "' ");
                    strSql.AppendLine("           , CONVERT(VARCHAR(20),GETDATE(),20) ");
                    strSql.AppendLine("           , "+ FmMainToolBar2.UserID + " ); ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //foreach (KeyValuePair<string, string> param in dicParams)
                    //{
                    //    cmd.Parameters.AddWithValue(param.Key, param.Value);
                    //}
                    cmd.ExecuteNonQuery();

                    //dicParams.Clear();
                    //dicParams.Add("@TDATE", sTDate);
                    //dicParams.Add("@ACCOD1", sAcCod1); //외상매입금, 미지급금에 대한 계정코드
                    //dicParams.Add("@ACCOD2", sAcCod2); //상대계정, 보통예금이나 현금에 대한 부분
                    //dicParams.Add("@CVCOD1", sCvCod); //외상매입금, 미지급금에 대한 거래처코드
                    //dicParams.Add("@CVCOD2", sActNo); //상대계정, 보통예금이나 현금에 대한 거래처코드
                    //dicParams.Add("@ATEXT", sAText); //적요
                    //dicParams.Add("@REF1", sSlipNo); //SUGMF의 PK값
                    //dicParams.Add("@ACAMT", dTrSum.ToString()); //대변 혼돈금지
                    //dicParams.Add("@ADAMT", dTrSum.ToString()); //차변 혼돈금지
                    //dicParams.Add("@USRCD", FmMainToolBar2.UserID); //유저코드
                    //dicParams.Add("@RK", sRk); //비고

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
                    //strSql.AppendLine("     , ( SELECT LPAD(CAST(IFNULL(MAX(Z1.SEQNO), 5000) AS DOUBLE) + 1, 4, '0') FROM ACTRAN Z1 WHERE Z1.TDATE = DATE_FORMAT(@TDATE, '%Y%m%d') AND Z1.ATGUB = '3' ) AS SEQNO ");
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
                    #endregion

                    strSql.AppendLine("SELECT '"+ sTDate.Replace("-","") + "' AS TDATE                                                                                ");
                    strSql.AppendLine("     , '3' AS ATGUB --대체전표                                                                            ");
                    strSql.AppendLine("     , (SELECT REPLICATE(0, 4 - LEN(ISNULL(MAX(Z1.SEQNO), 5000) + 1)) + ISNULL(MAX(Z1.SEQNO), 5000) + 1   ");
                    strSql.AppendLine("           FROM ACTRAN Z1                                                                                 ");
                    strSql.AppendLine("          WHERE Z1.TDATE = '" + sTDate.Replace("-", "") + "'                                                                     ");
                    strSql.AppendLine("            AND Z1.ATGUB = '3' ) AS SEQNO                                                                 ");
                    strSql.AppendLine("     , Y1.*                                                                                               ");
                    strSql.AppendLine("  FROM(                                                                                                   ");
                    strSql.AppendLine("         SELECT 0 AS LINNO--외상매입금 부분                                                               ");
                    strSql.AppendLine("              , '"+ sAcCod1 + "' AS ACCOD                                                                          ");
                    strSql.AppendLine("              , (SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = '"+ sAcCod1 + "') AS ACNAM                        ");
                    strSql.AppendLine("            , "+ sCvCod + " AS CVCOD                                                            ");
                    strSql.AppendLine("            , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = "+ sCvCod + " ) AS CVNAM         ");
                    strSql.AppendLine("            , '"+ sAText + "' AS ATEXT                                                                             ");
                    strSql.AppendLine("            , 0 AS ADAMT --외상매입금, 미지급금(ADAMT는 차변임!!! )                                       ");
                    strSql.AppendLine("              , "+ dTrSum + " AS ACAMT--외상매입금, 미지급금(ACAMT는 대변임!!! )                                 ");
                    strSql.AppendLine("              , (SELECT X2.DEPT_CD                                                                        ");
                    strSql.AppendLine("                   FROM ZUSRLST X1                                                                        ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                             ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                            ");
                    strSql.AppendLine("                   WHERE USRCD = "+ FmMainToolBar2.UserID + " ) AS ADPCD                                                        ");
                    strSql.AppendLine("              , ( SELECT X3.DEPT_NM                                                                       ");
                    strSql.AppendLine("                    FROM ZUSRLST X1                                                                       ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                             ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                            ");
                    strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3                                                              ");
                    strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD                                                          ");
                    strSql.AppendLine("                   WHERE USRCD = " + FmMainToolBar2.UserID + " ) AS ADPNM                                                        ");
                    strSql.AppendLine("              , '"+ sSlipNo + "' AS REF1 --SUGMF 테이블의 PK값                                                       ");
                    strSql.AppendLine("              , 'N' AS APVYN                                                                              ");
                    strSql.AppendLine("              , 'D01' AS AAUTO --자동전표 구분 HARDCODING                                                 ");
                    strSql.AppendLine("              , NULL AS ADATE                                                                             ");
                    strSql.AppendLine("              , NULL AS AUSER                                                                             ");
                    strSql.AppendLine("              , '"+ sRk + "' AS RK                                                                                 ");
                    strSql.AppendLine("              , " + FmMainToolBar2.UserID + " AS CUSER                                                                           ");
                    strSql.AppendLine("              , GETDATE() AS CDATE                                                                        ");
                    strSql.AppendLine("          UNION ALL                                                                                       ");
                    strSql.AppendLine("         SELECT 1 AS LINNO --상대계정 부분                                                                ");
                    strSql.AppendLine("              , '"+ sAcCod2 + "' AS ACCOD                                                            ");
                    strSql.AppendLine("              , ( SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = '"+ sAcCod2 + "' ) AS ACNAM                      ");
                    strSql.AppendLine("              , "+ sActNo + " AS CVCOD                                                          ");
                    strSql.AppendLine("              , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = "+ sActNo + " ) AS CVNAM       ");
                    strSql.AppendLine("              , '"+ sAText + "' AS ATEXT                                                                           ");
                    strSql.AppendLine("              , "+ dTrSum + " AS ADAMT --상대계정(ADAMT는 차변임!!! )                                            ");
                    strSql.AppendLine("              , 0 AS ACAMT --상대계정(ADAMT는 차변임!!! )                                                 ");
                    strSql.AppendLine("              , (SELECT X2.DEPT_CD                                                                        ");
                    strSql.AppendLine("                   FROM ZUSRLST X1                                                                        ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                             ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                            ");
                    strSql.AppendLine("                   WHERE USRCD = " + FmMainToolBar2.UserID + " ) AS ADPCD                                                        ");
                    strSql.AppendLine("              , ( SELECT X3.DEPT_NM                                                                       ");
                    strSql.AppendLine("                    FROM ZUSRLST X1                                                                       ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                             ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                            ");
                    strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3                                                              ");
                    strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD                                                          ");
                    strSql.AppendLine("                   WHERE USRCD = " + FmMainToolBar2.UserID + " ) AS ADPNM                                                        ");
                    strSql.AppendLine("              , '"+ sSlipNo + "' AS REF1 --SUGMF 테이블의 PK값                                                       ");
                    strSql.AppendLine("              , 'N' AS APVYN                                                                              ");
                    strSql.AppendLine("              , 'D01' AS AAUTO --자동전표 구분                                                            ");
                    strSql.AppendLine("              , NULL AS ADATE                                                                             ");
                    strSql.AppendLine("              , NULL AS AUSER                                                                             ");
                    strSql.AppendLine("              , '" + sRk + "' AS RK                                                                                 ");
                    strSql.AppendLine("              , " + FmMainToolBar2.UserID + " AS CUSER                                                                           ");
                    strSql.AppendLine("              , GETDATE() AS CDATE                                                                        ");
                    strSql.AppendLine("       ) Y1                                                                                               ");
                                                                                                                                                  
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //foreach (KeyValuePair<string, string> param in dicParams)
                    //{
                    //    cmd.Parameters.AddWithValue(param.Key, param.Value);
                    //}
                    cmd.ExecuteNonQuery();
                }
                //지불번호가 존재하지 않는 경우
                //전표 INSERT
                else
                {
                    //dicParams.Add("@SLIPNO", sSlipNo);
                    //dicParams.Add("@TDATE", sTDate);
                    //dicParams.Add("@JGUBN", sJGbun);
                    //dicParams.Add("@BILGU", sBilGu);
                    //dicParams.Add("@CVCOD", sCvCod);
                    //dicParams.Add("@ACTNO", sActNo);
                    //dicParams.Add("@TRSUM", dTrSum.ToString());
                    //dicParams.Add("@RK", sRk);
                    //dicParams.Add("@USRCD", FmMainToolBar2.UserID);

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE SUGMF ");
                    strSql.AppendLine("    SET TDATE = '"+ sTDate + "' ");
                    strSql.AppendLine("      , JGUBN = '"+ sJGbun + "' ");
                    strSql.AppendLine("      , BILGU = '"+ sBilGu + "' ");
                    strSql.AppendLine("      , CVCOD = "+ sCvCod + " ");
                    strSql.AppendLine("      , ACTNO = "+ sActNo + " ");
                    strSql.AppendLine("      , TRSUM = "+ dTrSum + " ");
                    strSql.AppendLine("      , RK = '"+ sRk + "' ");
                    strSql.AppendLine("      , MDATE = GETDATE() ");
                    strSql.AppendLine("      , MUSER = "+ FmMainToolBar2.UserID + " ");
                    strSql.AppendLine("  WHERE SLIPNO = '"+ sSlipNo + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //foreach (KeyValuePair<string, string> param in dicParams)
                    //{
                    //    cmd.Parameters.AddWithValue(param.Key, param.Value);
                    //}
                    cmd.ExecuteNonQuery();


                    /*
                     * 전표를 업데이트 하지 않고 DELETE 후 INSERT
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
                    string sRstTDate = dt.Rows[0]["TDATE"]?.ToString();
                    string sAtGub = dt.Rows[0]["ATGUB"]?.ToString();
                    string sSeqNo = dt.Rows[0]["SEQNO"]?.ToString();

                    if(string.IsNullOrEmpty(sRstTDate) || string.IsNullOrEmpty(sAtGub) || string.IsNullOrEmpty(sSeqNo))
                    {
                        throw new Exception("업데이트 중 전표정보가 없습니다.\r\n관리자에게 문의하세요.");
                    }

                    //dicParams.Clear();
                    //dicParams.Add("@TDATE", sRstTDate);
                    //dicParams.Add("@ATGUB", sAtGub);
                    //dicParams.Add("@SEQNO", sSeqNo);

                    strSql.Clear();
                    strSql.AppendLine(" DELETE FROM ACTRAN ");
                    strSql.AppendLine("       WHERE TDATE = '"+ sRstTDate + "' ");
                    strSql.AppendLine("         AND ATGUB = '"+ sAtGub + "' ");
                    strSql.AppendLine("         AND SEQNO = '"+ sSeqNo + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    //cmd.Parameters.Clear();
                    //foreach (KeyValuePair<string, string> param in dicParams)
                    //{
                    //    cmd.Parameters.AddWithValue(param.Key, param.Value);
                    //}
                    cmd.ExecuteNonQuery();

                    dicParams.Clear();
                    dicParams.Add("@TDATE", sTDate);
                    dicParams.Add("@ACCOD1", sAcCod1); //외상매입금, 미지급금에 대한 계정코드
                    dicParams.Add("@ACCOD2", sAcCod2); //상대계정, 보통예금이나 현금에 대한 부분
                    dicParams.Add("@CVCOD1", sCvCod); //외상매입금, 미지급금에 대한 거래처코드
                    dicParams.Add("@CVCOD2", sActNo); //상대계정, 보통예금이나 현금에 대한 거래처코드
                    dicParams.Add("@ATEXT", sAText); //적요
                    dicParams.Add("@REF1", sSlipNo); //SUGMF의 PK값
                    dicParams.Add("@ACAMT", dTrSum.ToString()); //대변 혼돈금지
                    dicParams.Add("@ADAMT", dTrSum.ToString()); //차변 혼돈금지
                    dicParams.Add("@USRCD", FmMainToolBar2.UserID); //유저코드
                    dicParams.Add("@RK", sRk); //비고

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
                    //strSql.AppendLine("     , ( SELECT LPAD(CAST(IFNULL(MAX(Z1.SEQNO), 5000) AS DOUBLE) + 1, 4, '0') FROM ACTRAN Z1 WHERE Z1.TDATE = DATE_FORMAT(@TDATE, '%Y%m%d') AND Z1.ATGUB = '3' ) AS SEQNO ");
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
                    #endregion

                    strSql.AppendLine("SELECT '" + sTDate.Replace("-", "") + "' AS TDATE                                                           ");
                    strSql.AppendLine("     , '3' AS ATGUB --대체전표                                                                          ");
                    strSql.AppendLine("     , (SELECT REPLICATE(0, 4 - LEN(ISNULL(MAX(Z1.SEQNO), 5000) + 1)) + ISNULL(MAX(Z1.SEQNO), 5000) + 1 ");
                    strSql.AppendLine("           FROM ACTRAN Z1                                                                               ");
                    strSql.AppendLine("          WHERE Z1.TDATE = '" + sTDate.Replace("-", "") + "'                                            ");
                    strSql.AppendLine("            AND Z1.ATGUB = '3' ) AS SEQNO                                                               ");
                    strSql.AppendLine("     , Y1.*                                                                                             ");
                    strSql.AppendLine("  FROM(                                                                                                 ");
                    strSql.AppendLine("         SELECT 0 AS LINNO--외상매입금 부분                                                             ");
                    strSql.AppendLine("              , '"+ sAcCod1 + "' AS ACCOD                                                          ");
                    strSql.AppendLine("              , (SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = '" + sAcCod1 + "') AS ACNAM                      ");
                    strSql.AppendLine("            , "+ sCvCod + " AS CVCOD                                                          ");
                    strSql.AppendLine("            , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = "+ sCvCod + " ) AS CVNAM       ");
                    strSql.AppendLine("            , '"+ sAText + "' AS ATEXT                                                                           ");
                    strSql.AppendLine("            , 0 AS ADAMT --외상매입금, 미지급금(ADAMT는 차변임!!! )                                     ");
                    strSql.AppendLine("              , "+ dTrSum + " AS ACAMT--외상매입금, 미지급금(ACAMT는 대변임!!! )                               ");
                    strSql.AppendLine("              , (SELECT X2.DEPT_CD                                                                      ");
                    strSql.AppendLine("                   FROM ZUSRLST X1                                                                      ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                          ");
                    strSql.AppendLine("                   WHERE USRCD = "+ FmMainToolBar2.UserID + " ) AS ADPCD                                                      ");
                    strSql.AppendLine("              , ( SELECT X3.DEPT_NM                                                                     ");
                    strSql.AppendLine("                    FROM ZUSRLST X1                                                                     ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                          ");
                    strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3                                                            ");
                    strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD                                                        ");
                    strSql.AppendLine("                   WHERE USRCD = "+ FmMainToolBar2.UserID + " ) AS ADPNM                                                      ");
                    strSql.AppendLine("              , '"+ sSlipNo + "' AS REF1 --SUGMF 테이블의 PK값                                                     ");
                    strSql.AppendLine("              , 'N' AS APVYN                                                                            ");
                    strSql.AppendLine("              , 'D01' AS AAUTO --자동전표 구분 HARDCODING                                               ");
                    strSql.AppendLine("              , NULL AS ADATE                                                                           ");
                    strSql.AppendLine("              , NULL AS AUSER                                                                           ");
                    strSql.AppendLine("              , '"+ sRk + "' AS RK                                                                               ");
                    strSql.AppendLine("              , "+ FmMainToolBar2.UserID + " AS CUSER                                                                         ");
                    strSql.AppendLine("              , GETDATE() AS CDATE                                                                      ");
                    strSql.AppendLine("          UNION ALL                                                                                     ");
                    strSql.AppendLine("         SELECT 1 AS LINNO --상대계정 부분                                                              ");
                    strSql.AppendLine("              , '"+ sAcCod2 + "' AS ACCOD                                                          ");
                    strSql.AppendLine("              , ( SELECT X1.ACNAM FROM ACMSTF X1 WHERE X1.ACCOD = '"+ sAcCod2 + "' ) AS ACNAM                    ");
                    strSql.AppendLine("              , "+ sActNo + " AS CVCOD                                                        ");
                    strSql.AppendLine("              , ( SELECT X1.DEALER_NM FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = "+ sActNo + " ) AS CVNAM     ");
                    strSql.AppendLine("              , '"+ sAText + "' AS ATEXT                                                                         ");
                    strSql.AppendLine("              , "+ dTrSum + " AS ADAMT --상대계정(ADAMT는 차변임!!! )                                          ");
                    strSql.AppendLine("              , 0 AS ACAMT --상대계정(ADAMT는 차변임!!! )                                               ");
                    strSql.AppendLine("              , (SELECT X2.DEPT_CD                                                                      ");
                    strSql.AppendLine("                   FROM ZUSRLST X1                                                                      ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                          ");
                    strSql.AppendLine("                   WHERE USRCD = "+ FmMainToolBar2.UserID + " ) AS ADPCD                                                      ");
                    strSql.AppendLine("              , ( SELECT X3.DEPT_NM                                                                     ");
                    strSql.AppendLine("                    FROM ZUSRLST X1                                                                     ");
                    strSql.AppendLine("                    LEFT JOIN HR_EMP_BASIS X2                                                           ");
                    strSql.AppendLine("                      ON X1.INSANO = X2.EMP_ID                                                          ");
                    strSql.AppendLine("                    LEFT JOIN ACC_DEPT_CD X3                                                            ");
                    strSql.AppendLine("                      ON X2.DEPT_CD = X3.DEPT_CD                                                        ");
                    strSql.AppendLine("                   WHERE USRCD = "+ FmMainToolBar2.UserID + " ) AS ADPNM                                                      ");
                    strSql.AppendLine("              , '"+ sSlipNo + "' AS REF1 --SUGMF 테이블의 PK값                                                     ");
                    strSql.AppendLine("              , 'N' AS APVYN                                                                            ");
                    strSql.AppendLine("              , 'D01' AS AAUTO --자동전표 구분                                                          ");
                    strSql.AppendLine("              , NULL AS ADATE                                                                           ");
                    strSql.AppendLine("              , NULL AS AUSER                                                                           ");
                    strSql.AppendLine("              , '"+ sRk + "' AS RK                                                                               ");
                    strSql.AppendLine("              , "+ FmMainToolBar2.UserID + " AS CUSER                                                                         ");
                    strSql.AppendLine("              , GETDATE() AS CDATE                                                                      ");
                    strSql.AppendLine("       ) Y1                                                                                             ");

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
                XtraMessageBox.Show("저장을 완료했습니다.");

                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void AC16001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
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

            //if (sNullYn.Equals("Y"))
            //{
            //    strSql.AppendLine(" SELECT '****' AS CD ");
            //    strSql.AppendLine("      , '선택' AS NM ");
            //    strSql.AppendLine("  UNION ALL ");
            //}

            if (sGb.Equals("1")) //지급구분
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine("  WHERE CD_GB = 'PAY_METHOD' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
            }
            else if (sGb.Equals("2")) //결제대상
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine("  WHERE CD_GB = 'AC16001_01' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
                strSql.AppendLine("    AND COM_CD IN ('1', '2') ");
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

        public DataRow DrDealerInfo;
        private void BtnEditCvNam_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            AC02001F03 frm = new AC02001F03();
            frm.P_AC16001F01 = this;
            frm.ACNT_YN = "N";
            if(frm.ShowDialog() == DialogResult.OK)
            {
                BtnEditCvNam.EditValue = DrDealerInfo["DEALER_NM"];
                TxtCvCod.EditValue = DrDealerInfo["DEALER_CD"];
                DrDealerInfo = null;
            }
        }

        private void BtnEditCvNam_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sDealerCd = btnEdit.EditValue?.ToString();
                if (!string.IsNullOrEmpty(sDealerCd))
                {
                    DataTable dt = GetDealerInfo(sDealerCd);
                    if (dt.Rows.Count == 1)
                    {
                        BtnEditCvNam.EditValue = dt.Rows[0]["DEALER_NM"];
                        TxtCvCod.EditValue = dt.Rows[0]["DEALER_CD"];
                    }
                    else
                    {
                        AC02001F03 frm = new AC02001F03();
                        frm.P_AC16001F01 = this;
                        frm.DealerCd = sDealerCd;
                        frm.ACNT_YN = "N";
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            BtnEditCvNam.EditValue = DrDealerInfo["DEALER_NM"];
                            TxtCvCod.EditValue = DrDealerInfo["DEALER_CD"];
                        }
                    }
                }
                else
                {
                    BtnEditCvNam.EditValue = null;
                    TxtCvCod.EditValue = null;
                }

                DrDealerInfo = null;
            }
        }

        private DataTable GetDealerInfo(string sDealerCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.DEALER_CD  ");
            strSql.AppendLine(" 	 , CAST(A.DEALER_NM AS VARCHAR) AS DEALER_NM ");
            strSql.AppendLine(" 	 , A.IDT_NO  ");
            strSql.AppendLine(" 	 , A.DEALER_GB  ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A  ");
            strSql.AppendLine("  WHERE (A.DEALER_CD LIKE '" + sDealerCd + "' ");
            strSql.AppendLine("     OR A.DEALER_NM LIKE '%" + sDealerCd + "%' ");
            strSql.AppendLine("     OR A.INITIAL_NM LIKE '%" + sDealerCd + "%' ) "); //거래처 초서검색 추가
            strSql.AppendLine("    AND A.EOB_YN = 'N' ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        private void BtnEditActNo_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            AC02001F03 frm = new AC02001F03();
            frm.P_AC16001F01 = this;
            frm.ACNT_YN = "Y";
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnEditActNo.EditValue = DrDealerInfo["DEALER_NM"];
                TxtAcntNo.EditValue = DrDealerInfo["DEALER_CD"];
                DrDealerInfo = null;
            }
        }

        private void BtnEditActNo_Leave(object sender, EventArgs e)
        {
            
        }

        private void BtnEditActNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sDealerCd = btnEdit.EditValue?.ToString();
                if (!string.IsNullOrEmpty(sDealerCd))
                {
                    DataTable dt = GetDealerInfo(sDealerCd);
                    if (dt.Rows.Count == 1)
                    {
                        BtnEditActNo.EditValue = dt.Rows[0]["DEALER_NM"];
                        TxtAcntNo.EditValue = dt.Rows[0]["DEALER_CD"];
                    }
                    else
                    {
                        AC02001F03 frm = new AC02001F03();
                        frm.P_AC16001F01 = this;
                        frm.DealerCd = sDealerCd;
                        frm.ACNT_YN = "Y";
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            BtnEditActNo.EditValue = DrDealerInfo["DEALER_NM"];
                            TxtAcntNo.EditValue = DrDealerInfo["DEALER_CD"];
                        }
                    }
                }
                else
                {
                    BtnEditActNo.EditValue = null;
                    TxtAcntNo.EditValue = null;
                }

                DrDealerInfo = null;
            }
        }
    }
}