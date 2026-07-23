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
using System.Collections;
/*
 * 작성일자 : 모름
 * 작성장 : 고혜성
 * ------------------------HISTORY-----------------------
 //* 1. 수정일자 : 2021-02-07
 //*    수정자 : 고혜성
 //*    수정내용 : 1) 외상대 지불 건만 나올 수 있도록 수정
 //*               -> 주의사항 : 현재 쿼리 수정은 외상대 지불이 포함된 건을 WHERE에서 LIKE로 가져오기 때문에 추후 수정필요
 //*                             -> 사유 : 기존에는 지불관리에서 올라온 데이터를 기준하면 아무 문제가 없으나 전표관리에서 외상대 지불을 입력하는 경우가 있어
 //*                                       해당 쿼리를 그렇게 입력하였지만 
 *   2. 수정일자 : 2021-02-07
 *      수정자 : 고혜성
 *      수정내용 : 거래처초성검색 추가
 *      
 *   3. 수정일자 : 2021-02-15
 *      수정자 : 고혜성
 *      수정내용 : 엑셀 출력 시 금액부분이 Text화 되는 현상 발견 -> 쿼리문에서 Format 지우는 것으로 해결 (쿼리참조)
 *      
 *      
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2022-09-29
 * 수정자   : 정은영
 * ID       : #0001
 * 수정내용 : (현업요청)
 *            1. 미결제금액 승인된 건만 조회되도록 변경
 *            
 * 수정일자 : 2022-09-30
 * 수정자   : 정은영
 * ID       : #0002
 * 수정내용 : 1. 미지급액 및 잔액 안맞는부분 수정
 * 
 * 수정일자 : 2022-10-11
 * 수정자   : 정은영
 * ID       : #0003
 * 수정내용 : (현업요청)
 *            1. 승인안된것도 표시 되도록 변경 -> 승인된 건만 보이도록 변경(2차요청)
 *            2. 잔액 NULL 일때 0으로 변경 -> 미체크시 값이 틀림.
 *            3. 잔액 = 미결제금액-금액
 *            
 * 수정일자 : 2022-10-17
 * 수정자   : 정은영
 * ID       : #0004
 * 수정내용 : (현업요청)
 *            1. 지불 건 GROUP BY 제거
 *            
 * 수정일자 : 2022-10-25
 * 수정자   : 정은영
 * ID       : #0005
 * 수정내용 : 1. 미결제금액 중복 합산 되는 부분 수정
 */
namespace AccAdm
{
    public partial class AC17001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC17001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC17001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today; 
            DateEditTo.EditValue = DateTime.Today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr };
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
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            int iFindIdx = CboFindSbj.SelectedIndex;
            string sFindWord = TxtFindWord.EditValue?.ToString();
            int iJGubn = CboAcNam.SelectedIndex;

            if (string.IsNullOrEmpty(sYmdFrom))
            {
                XtraMessageBox.Show("전표일자를 입력하세요.");
                DateEditFrom.SelectAll();
                DateEditFrom.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("전표일자를 입력하세요.");
                DateEditTo.SelectAll();
                DateEditTo.Focus();
                return;
            }

            if(iFindIdx < 0)
            {
                XtraMessageBox.Show("찾을항목을 선택하세요.");
                CboFindSbj.Focus();
                return;
            }

            if(iJGubn < 0)
            {
                XtraMessageBox.Show("조회계정을 선택하세요.");
                CboAcNam.Focus();
                return;
            }

            //사업자번호일 경우 '-'를 공백처리
            if(iFindIdx == 1)
            {
                sFindWord.Replace("-", "");
                sFindWord.Replace(" ", "");
            }

            GetInfoNew();

            //Dictionary<string, string> dicParams = new Dictionary<string, string>();

            //dicParams.Add("DATE_F", sYmdFrom);
            //dicParams.Add("DATE_T", sYmdTo);
            //dicParams.Add("FIND_IDX", iFindIdx.ToString());
            //dicParams.Add("FIND_WORD", sFindWord);
            //dicParams.Add("AC_GB", iJGubn.ToString());

            //GridRetr.DataSource = GetInfo(dicParams);
            //if(GridViewRetr.RowCount == 0)
            //{
            //    DateEditFrom.SelectAll();
            //    DateEditFrom.Focus();
            //}
            //else if(GridViewRetr.RowCount > 0)
            //{
            //    GridViewRetr.Focus();
            //}
        }

        private void GetInfoNew()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                ArrayList arr = new ArrayList();

                arr.Add(DateTime.Now.ToString());
                arr.Add(string.Format("{0} ~ {1}", DateEditFrom.EditValue?.ToString().Substring(0, 10), DateEditTo.EditValue?.ToString().Substring(0, 10)));

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                int iFindIdx = CboFindSbj.SelectedIndex;
                string sFindWord = TxtFindWord.EditValue?.ToString();
                if (string.IsNullOrEmpty(TxtFindWord.EditValue?.ToString()))
                    sFindWord = "";
                int iJGubn = CboAcNam.SelectedIndex;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);
                dicParams.Add("FIND_IDX", iFindIdx.ToString());
                dicParams.Add("FIND_WORD", sFindWord);
                dicParams.Add("AC_GB", iJGubn.ToString());

                DataTable dt = GetInfo(dicParams, "");

                if (dt.Rows.Count == 0)
                {
                    string sMsg = string.Format("조회기간 : {0}" +
                        "\r\n해당기간 동안 발생된 지출내역이 존재하지 않습니다.", arr[1].ToString());

                    Cursor = Cursors.Default;
                    XtraMessageBox.Show(sMsg);
                    return;
                }

                int iTotAmt = 0;
                foreach (DataRow row in dt.Rows)
                {
                    string sVal = row["TRSUM"]?.ToString();
                    int i = string.IsNullOrEmpty(sVal.Replace(",", "")) ? 0 : Convert.ToInt32(sVal.Replace(",", ""));
                    iTotAmt += i;
                }

                #region 2022-10-11 이후 주석 #0003
                //for (int i = 0; i < dt.Rows.Count; i++)
                //{
                //    string sCvCod = dt.Rows[i]["CVCOD"]?.ToString();
                //    string sAcCod = dt.Rows[i]["ACCOD"]?.ToString();
                //    string sTrSum = dt.Rows[i]["TRSUM"]?.ToString();

                //    double dTrSum = string.IsNullOrEmpty(sTrSum) ? 0 : Convert.ToDouble(sTrSum);

                //    dicParams.Clear();
                //    dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8));
                //    dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8));
                //    dicParams.Add("ACCOD", sAcCod);
                //    dicParams.Add("CVCOD", sCvCod);

                //    DataTable dtTemp = GetInfo2(dicParams);
                //    if (dtTemp.Rows.Count > 0)
                //    {
                //        string sOfsBlc = dtTemp.Rows[0]["OFS_BLC"]?.ToString();
                //        double dOfsBls = string.IsNullOrEmpty(sOfsBlc) ? 0 : Convert.ToDouble(sOfsBlc);

                //        //#0003
                //        double dAmt = dOfsBls + dTrSum;
                //        dt.Rows[i]["AMT"] = dAmt;
                //        dt.Rows[i]["DIFF"] = dAmt - dTrSum;

                //        //#0002
                //        //double dAmt = 0;
                //        //if(dOfsBls != 0)
                //        //{
                //        //    dAmt = dOfsBls - dTrSum;
                //        //}
                //        //
                //        //dt.Rows[i]["AMT"] = dOfsBls;
                //        //dt.Rows[i]["DIFF"] = dAmt;
                //    }
                //}
                #endregion

                Cursor = Cursors.Default;
                GridRetr.DataSource = dt;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private DataTable GetInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.SLIPNO ");
            strSql.AppendLine("      , A.CVCOD ");
            strSql.AppendLine("      , B.DEALER_NM ");
            strSql.AppendLine("      , B.IDT_NO ");
            strSql.AppendLine("      , B1.EMP_NM  ");
            strSql.AppendLine("      , A.TDATE ");
            strSql.AppendLine("      , (SELECT X1.ATEXT FROM ACTRAN X1 WHERE X1.REF1 = A.SLIPNO AND X1.AAUTO = 'D01' LIMIT 1) AS RK ");
            strSql.AppendLine("      , 0 AS OFS_BLC ");
            strSql.AppendLine("      , A.TRSUM ");
            strSql.AppendLine("      , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK ");
            strSql.AppendLine("      , B.BANK_ACNT_NO  ");
            strSql.AppendLine("      , B.ACNT_HOLDER  ");
            strSql.AppendLine("   FROM SUGMF A ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B ");
            strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1 ");
            strSql.AppendLine("     ON B.CHRG_ID = B1.EMP_ID ");
            strSql.AppendLine("   LEFT JOIN COM_BASE_CD C ");
            strSql.AppendLine("     ON B.BANK_CD = C.COM_CD ");
            strSql.AppendLine("    AND C.CD_GB = 'BANK_CD' ");
            strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "' ");
            strSql.AppendLine("    AND (('" + dicParams["AC_GB"] + "' = '0' AND A.JGUBN = '1')  ");
            strSql.AppendLine("         OR ");
            strSql.AppendLine("         ('" + dicParams["AC_GB"] + "' = '1' AND A.JGUBN = '2' ))  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '2' AND A.JGUBN = '2' ))  ");
            //strSql.AppendLine("    AND ((@FIND_WORD = '' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '0' AND B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '2' AND B1.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))  ");

            DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);


            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private DataTable GetInfo(Dictionary<string, string> dicParams, string sGb)
        {
            StringBuilder strSql = new StringBuilder();

            //#0004
            strSql.AppendLine(" WITH TEMP1 AS(                                                          ");
            strSql.AppendLine("     SELECT A.SLIPNO                                                     ");
            strSql.AppendLine("          , A.CVCOD                                                      ");
            strSql.AppendLine("          , A.JGUBN                                                      ");
            strSql.AppendLine("          , CASE WHEN A.JGUBN = '1' THEN '0251' ELSE '0253' END AS ACCOD ");
            strSql.AppendLine("          , A.TDATE                                                      ");
            strSql.AppendLine("          , ISNULL(A.TRSUM, 0) AS TRSUM                                  ");
            strSql.AppendLine("       FROM SUGMF A                                                      ");
            strSql.AppendLine("      WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T                ");
            strSql.AppendLine(" ),INFO AS(                                                              ");
            strSql.AppendLine("      SELECT X1.CVCOD                                                    ");
            strSql.AppendLine("           , X1.ACCOD                                                    ");
            strSql.AppendLine("           , SUM(X1.JAMT) AS JAMT                                        ");
            strSql.AppendLine("        FROM (                                                           ");
            strSql.AppendLine("               SELECT A1.ACCOD                                           ");
            strSql.AppendLine("                    , A1.CVCOD                                           ");
            strSql.AppendLine("                    , SUM(ISNULL(A1.ACDRJN,0))+SUM(ISNULL(A1.ACCRJN, 0)) AS JAMT");
            strSql.AppendLine("                  FROM ACJANF A1                                                ");
            strSql.AppendLine("                 WHERE A1.ACYEAR = SUBSTRING(@DATE_F, 1, 4)                ");
            strSql.AppendLine("                 GROUP BY A1.CVCOD, A1.ACCOD                                    ");
            strSql.AppendLine("                 UNION ALL                                                      ");
            strSql.AppendLine("                SELECT A1.ACCOD                                                 ");
            strSql.AppendLine("                     , A1.CVCOD                                                 ");
            strSql.AppendLine("                     , ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) END), 0) AS JAMT");
            strSql.AppendLine("                  FROM ACTRAN A1            ");
            strSql.AppendLine("                  LEFT JOIN ACMSTF B1       ");
            strSql.AppendLine("                    ON A1.ACCOD = B1.ACCOD  ");
            strSql.AppendLine("                 WHERE A1.TDATE BETWEEN SUBSTRING(@DATE_F, 1, 4) + '0101' AND REPLACE(DATEADD(DAY, -1, CONVERT(DATE, @DATE_F)),'-','')");
            strSql.AppendLine("                   AND A1.APVYN = 'Y'                ");
            strSql.AppendLine("                 GROUP BY A1.CVCOD, A1.ACCOD         ");
            strSql.AppendLine("              ) X1                                   ");
            strSql.AppendLine("        GROUP BY X1.CVCOD, X1.ACCOD                  ");
            strSql.AppendLine("  ), INFO2 AS(                                       ");
            strSql.AppendLine("       SELECT X1.CVCOD                               ");
            strSql.AppendLine("            , X1.ACCOD                               ");
            strSql.AppendLine("            , MAX(X1.ATEXT) AS RMK                   ");
            strSql.AppendLine("            , SUM(ISNULL(X1.ADAMT, 0)) AS OCR_AMT    ");
            strSql.AppendLine("            , SUM(ISNULL(X1.ACAMT, 0)) AS CPT_OFS_AMT");
            strSql.AppendLine("            , 0 AS OFS_AMT                           ");
            strSql.AppendLine("         FROM ACTRAN X1                              ");
            strSql.AppendLine("         LEFT JOIN ACC_DEALER_CD X2                  ");
            strSql.AppendLine("           ON X1.CVCOD = X2.DEALER_CD                ");
            strSql.AppendLine("        WHERE X1.TDATE BETWEEN REPLACE(@DATE_F, '-', '') AND REPLACE(@DATE_T,  '-', '')");
            strSql.AppendLine("          AND X1.APVYN = 'Y'                   ");
            strSql.AppendLine("        GROUP BY X1.CVCOD, X1.ACCOD ");
            strSql.AppendLine("  ), TEMP2 AS(                                 ");
            strSql.AppendLine("       SELECT ROW_NUMBER() OVER(PARTITION BY A.CVCOD,A.ACCOD ORDER BY A.TDATE) AS ROWNUM");
            strSql.AppendLine("            , A.SLIPNO                                                           ");
 	        strSql.AppendLine("           , A.CVCOD                                                             ");
	        strSql.AppendLine("           , B.DEALER_NM                                                         ");
            strSql.AppendLine("           , A.ACCOD                                                             ");
            strSql.AppendLine("           , CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-', SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS IDT_NO");
            strSql.AppendLine("           , B1.EMP_NM AS CHRG_NM         ");
	        strSql.AppendLine("           , A.TDATE                      ");
	        strSql.AppendLine("           , E.RMK                        ");
	        strSql.AppendLine("           , ISNULL(A.TRSUM, 0) AS TRSUM  ");
            strSql.AppendLine("           , ISNULL(D.JAMT, 0) AS AMT     ");
            strSql.AppendLine("           , ISNULL(D.JAMT, 0) -ISNULL(A.TRSUM, 0) AS DIFF");
            strSql.AppendLine("          , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK");
            strSql.AppendLine("          , B.BANK_ACNT_NO         ");
	        strSql.AppendLine("           , B.ACNT_HOLDER         ");
            strSql.AppendLine("        FROM TEMP1 A               ");
            strSql.AppendLine("        LEFT JOIN ACC_DEALER_CD B  ");
            strSql.AppendLine("          ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendLine("        LEFT JOIN HR_EMP_BASIS B1  ");
            strSql.AppendLine("          ON B.CHRG_ID = B1.EMP_ID ");
            strSql.AppendLine("        LEFT JOIN COM_BASE_CD C    ");
            strSql.AppendLine("          ON B.BANK_CD = C.COM_CD  ");
            strSql.AppendLine("         AND C.CD_GB = 'BANK_CD'   ");
            strSql.AppendLine("        LEFT JOIN INFO D           ");
            strSql.AppendLine("          ON A.CVCOD = D.CVCOD     ");
            strSql.AppendLine("         AND D.ACCOD = A.ACCOD     ");
            strSql.AppendLine("        LEFT JOIN INFO2 E          ");
            strSql.AppendLine("          ON A.CVCOD = E.CVCOD     ");
            strSql.AppendLine("         AND E.ACCOD = A.ACCOD     ");
            strSql.AppendLine(" WHERE 1=1");
            strSql.AppendLine("   AND ((@AC_GB = '0' AND 1 = 1)  ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@AC_GB = '1' AND A.JGUBN = '1' )  ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@AC_GB = '2' AND A.JGUBN = '2' ))  ");
            strSql.AppendLine("   AND ((@FIND_WORD = '' AND 1 = 1)  ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@FIND_IDX = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))  ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        (@FIND_IDX = '2' AND B1.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))  ");
            strSql.AppendLine("  )                                                                          ");
            strSql.AppendLine("                                                                             ");
            strSql.AppendLine("  SELECT Z1.ROWNUM                                                           ");
            strSql.AppendLine("       , Z1.SLIPNO                                                           ");
            strSql.AppendLine("       , Z1.CVCOD                                                            ");
            strSql.AppendLine("       , Z1.DEALER_NM                                                        ");
            strSql.AppendLine("       , Z1.ACCOD                                                            ");
            strSql.AppendLine("       , Z1.IDT_NO                                                           ");
            strSql.AppendLine("       , Z1.CHRG_NM                                                          ");
            strSql.AppendLine("       , Z1.TDATE                                                            ");
            strSql.AppendLine("       , Z1.RMK                                                              ");
            strSql.AppendLine("       , Z1.TRSUM                                                            ");
            strSql.AppendLine("       , CASE WHEN ROWNUM = 1 THEN Z1.JAN                                    ");
            strSql.AppendLine("              ELSE ISNULL(LAG(DIFF) OVER(ORDER BY Z1.CVCOD, Z1.SLIPNO), 0) END AS AMT");
            strSql.AppendLine("       , Z1.DIFF          ");
            strSql.AppendLine("       , Z1.BANK          ");
            strSql.AppendLine("       , Z1.BANK_ACNT_NO  ");
            strSql.AppendLine("       , Z1.ACNT_HOLDER   ");
            strSql.AppendLine("    FROM(                 ");
            strSql.AppendLine("        SELECT X1.ROWNUM  ");
            strSql.AppendLine("       , X1.SLIPNO        ");
            strSql.AppendLine("       , X1.CVCOD         ");
            strSql.AppendLine("       , X1.DEALER_NM     ");
            strSql.AppendLine("       , X1.ACCOD         ");
            strSql.AppendLine("       , X1.IDT_NO        ");
            strSql.AppendLine("       , X1.CHRG_NM       ");
            strSql.AppendLine("       , X1.TDATE         ");
            strSql.AppendLine("       , X1.RMK           ");
            strSql.AppendLine("       , X1.TRSUM         ");
            strSql.AppendLine("       , X1.AMT           ");
            strSql.AppendLine("       , X1.BANK          ");
            strSql.AppendLine("       , X1.BANK_ACNT_NO  ");
            strSql.AppendLine("       , X1.ACNT_HOLDER   ");
            strSql.AppendLine("       , X1.JAN           ");
            strSql.AppendLine("       , SUM(X1.JAN - X1.TRSUM) OVER(PARTITION BY X1.CVCOD ORDER BY X1.CVCOD, X1.SLIPNO) AS DIFF");
            strSql.AppendLine("    FROM(SELECT ROWNUM          ");
            strSql.AppendLine("                 , SLIPNO       ");
            strSql.AppendLine("                 , CVCOD        ");
            strSql.AppendLine("                 , DEALER_NM    ");
            strSql.AppendLine("                 , ACCOD        ");
            strSql.AppendLine("                 , IDT_NO       ");
            strSql.AppendLine("                 , CHRG_NM      ");
            strSql.AppendLine("                 , TDATE        ");
            strSql.AppendLine("                 , RMK          ");
            strSql.AppendLine("                 , TRSUM        ");
            strSql.AppendLine("                 , AMT          ");
            strSql.AppendLine("                 , BANK         ");
            strSql.AppendLine("                 , BANK_ACNT_NO ");
            strSql.AppendLine("                 , ACNT_HOLDER  ");
            strSql.AppendLine("                 , CASE WHEN ROWNUM = 1 THEN AMT ELSE 0 END JAN");
            strSql.AppendLine("              FROM TEMP2) X1");
            strSql.AppendLine("    )Z1                     ");

            #region 2022-10-17 이전
            //#0003
            //   strSql.AppendLine(" WITH TEMP1 AS(                                                          ");
            //   strSql.AppendLine("     SELECT A.CVCOD                                                      ");
            //   strSql.AppendLine("          , A.JGUBN");
            //   strSql.AppendLine("          , CASE WHEN A.JGUBN = '1' THEN '0251' ELSE '0253' END AS ACCOD ");
            //   strSql.AppendLine("          , A.TDATE                                                      ");
            //   strSql.AppendLine("          , SUM(ISNULL(A.TRSUM, 0)) AS TRSUM                             ");
            //   strSql.AppendLine("       FROM SUGMF A                                                      ");
            //   strSql.AppendLine("      WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T                ");
            //   strSql.AppendLine("      GROUP BY A.CVCOD, A.JGUBN, A.TDATE                                 ");
            //   strSql.AppendLine(" ),INFO AS(                                                              ");
            //   strSql.AppendLine("      SELECT X1.CVCOD                                                    ");
            //   strSql.AppendLine("           , X1.ACCOD                                                    ");
            //   strSql.AppendLine("           , SUM(X1.JAMT) AS JAMT                                        ");
            //   strSql.AppendLine("        FROM (                                                           ");
            //   strSql.AppendLine("               SELECT A1.ACCOD                                           ");
            //   strSql.AppendLine("                    , A1.CVCOD                                           ");
            //   strSql.AppendLine("                    , SUM(ISNULL(A1.ACDRJN,0))+SUM(ISNULL(A1.ACCRJN, 0)) AS JAMT ");
            //   strSql.AppendLine("                 FROM ACJANF A1                                                  ");
            //   strSql.AppendLine("                WHERE A1.ACYEAR = SUBSTRING(@DATE_F, 1, 4)                  ");
            //   strSql.AppendLine("                GROUP BY A1.CVCOD, A1.ACCOD                                      ");
            //   strSql.AppendLine("                UNION ALL                                                        ");
            //   strSql.AppendLine("               SELECT A1.ACCOD                                                   ");
            //   strSql.AppendLine("                    , A1.CVCOD                                                   ");
            //   strSql.AppendLine("                    , ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) END), 0) AS JAMT");
            //   strSql.AppendLine("                 FROM ACTRAN A1           ");
            //   strSql.AppendLine("                 LEFT JOIN ACMSTF B1      ");
            //   strSql.AppendLine("                   ON A1.ACCOD = B1.ACCOD ");
            //   strSql.AppendLine("                WHERE A1.TDATE BETWEEN SUBSTRING(@DATE_F, 1, 4) + '0101' AND REPLACE(DATEADD(DAY, -1, CONVERT(DATE, @DATE_F)),'-','')");
            //   strSql.AppendLine("                  AND A1.APVYN = 'Y'                ");
            //   strSql.AppendLine("                GROUP BY A1.CVCOD, A1.ACCOD         ");
            //   strSql.AppendLine("             ) X1                                   ");
            //   strSql.AppendLine("       GROUP BY X1.CVCOD, X1.ACCOD                  ");
            //   strSql.AppendLine(" ), INFO2 AS(                                       ");
            //   strSql.AppendLine("      SELECT X1.CVCOD                               ");
            //   strSql.AppendLine("           , X1.ACCOD                               ");
            //   strSql.AppendLine("           , X1.TDATE");
            //   strSql.AppendLine("           , MAX(X1.ATEXT) AS RMK                   ");
            //   strSql.AppendLine("           , SUM(ISNULL(X1.ADAMT, 0)) AS OCR_AMT    ");
            //   strSql.AppendLine("           , SUM(ISNULL(X1.ACAMT, 0)) AS CPT_OFS_AMT");
            //   strSql.AppendLine("           , 0 AS OFS_AMT                           ");
            //   strSql.AppendLine("        FROM ACTRAN X1                              ");
            //   strSql.AppendLine("        LEFT JOIN ACC_DEALER_CD X2                  ");
            //   strSql.AppendLine("          ON X1.CVCOD = X2.DEALER_CD                ");
            //   strSql.AppendLine("       WHERE X1.TDATE BETWEEN REPLACE(@DATE_F, '-', '') AND REPLACE(@DATE_T,  '-', '')");
            //strSql.AppendLine("         AND X1.APVYN = 'Y'         ");
            //   strSql.AppendLine("       GROUP BY X1.CVCOD, X1.ACCOD, X1.TDATE    ");
            //   strSql.AppendLine(" )                                  ");
            //   strSql.AppendLine(" SELECT A.CVCOD                     ");
            //   strSql.AppendLine("      , B.DEALER_NM                 ");
            //   strSql.AppendLine("      , A.ACCOD                     ");
            //   strSql.AppendLine("      , CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-', SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS IDT_NO");
            //   strSql.AppendLine("      , B1.EMP_NM AS CHRG_NM        ");
            //   strSql.AppendLine("      , A.TDATE                     ");
            //   strSql.AppendLine("      , E.RMK                       ");
            //   strSql.AppendLine("      , ISNULL(A.TRSUM, 0) AS TRSUM ");
            //   strSql.AppendLine("      , ISNULL(D.JAMT, 0) AS AMT    ");
            //   strSql.AppendLine("      , ISNULL(D.JAMT, 0) - ISNULL(A.TRSUM, 0) AS DIFF");
            //   strSql.AppendLine("      , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK       ");
            //   strSql.AppendLine("      , B.BANK_ACNT_NO                                                            ");
            //   strSql.AppendLine("      , B.ACNT_HOLDER                                                             ");
            //   strSql.AppendLine("   FROM TEMP1 A                                                                   ");
            //   strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B                                                      ");
            //   strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD                                                     ");
            //   strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1                                                      ");
            //   strSql.AppendLine("     ON B.CHRG_ID = B1.EMP_ID                                                     ");
            //   strSql.AppendLine("   LEFT JOIN COM_BASE_CD C                                                        ");
            //   strSql.AppendLine("     ON B.BANK_CD = C.COM_CD                                                      ");
            //   strSql.AppendLine("    AND C.CD_GB = 'BANK_CD'                                                       ");
            //   strSql.AppendLine("   LEFT JOIN INFO D                                                               ");
            //   strSql.AppendLine("     ON A.CVCOD = D.CVCOD                                                         ");
            //   strSql.AppendLine("    AND D.ACCOD = A.ACCOD                                                         ");
            //   strSql.AppendLine("   LEFT JOIN INFO2 E                                                              ");
            //   strSql.AppendLine("     ON A.CVCOD = E.CVCOD                                                         ");
            //   strSql.AppendLine("    AND E.ACCOD = A.ACCOD                                                         ");
            //   strSql.AppendLine("    AND E.TDATE = A.TDATE");
            //   strSql.AppendLine(" WHERE 1=1");
            //   strSql.AppendLine("   AND ((@AC_GB = '0' AND 1 = 1)  ");
            //   strSql.AppendLine("        OR ");
            //   strSql.AppendLine("        (@AC_GB = '1' AND A.JGUBN = '1' )  ");
            //   strSql.AppendLine("        OR ");
            //   strSql.AppendLine("        (@AC_GB = '2' AND A.JGUBN = '2' ))  ");
            //   strSql.AppendLine("   AND ((@FIND_WORD = '' AND 1 = 1)  ");
            //   strSql.AppendLine("        OR ");
            //   strSql.AppendLine("        (@FIND_IDX = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))  ");
            //   strSql.AppendLine("        OR ");
            //   strSql.AppendLine("        (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //   strSql.AppendLine("        OR ");
            //   strSql.AppendLine("        (@FIND_IDX = '2' AND B1.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))  ");
            //   strSql.AppendLine("  ORDER BY A.TDATE, REPLACE(B.DEALER_NM, '(주)', '')");
            #endregion

            #region 2022-10-11 이전 
            /*
             * 2021-02-09 (현업요청)
             * 기준테이블을 전표테이블로 조회하여 수정했던 것에서 다시 SUGMF을 기준으로 조회
             * 현업에서 미지급 및 외상매입금의 경우 지불관리에서만 작성한다고 하여 다시 RollBack
             * 쿼리는 아래 region태그 중 '2021-01-13 이전쿼리 그대로 사용'
             */
            //strSql.Clear();
            //strSql.AppendLine("SELECT A.SLIPNO                                                                                                      ");
            //strSql.AppendLine("      , A.CVCOD                                                                                                      ");
            //strSql.AppendLine("      , B.DEALER_NM                                                                                                  ");
            //strSql.AppendLine("      , CASE WHEN A.JGUBN = '1' THEN '0251' ELSE '0253' END AS ACCOD                                                 ");
            //strSql.AppendLine("      , CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-', SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS IDT_NO  ");
            //strSql.AppendLine("      , B1.EMP_NM AS CHRG_NM                                                                                         ");
            //strSql.AppendLine("      , A.TDATE                                                                                                      ");
            //strSql.AppendLine("      , (SELECT TOP 1 X1.ATEXT FROM ACTRAN X1 WHERE X1.REF1 = A.SLIPNO AND X1.AAUTO = 'D01') AS RMK                  ");
            //strSql.AppendLine("      , 0 AS AMT                                                                                                     ");
            //strSql.AppendLine("      , A.TRSUM AS TRSUM                                                                                             ");
            //strSql.AppendLine("      , 0 AS DIFF                                                                                                    ");
            //strSql.AppendLine("      , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK                                          ");
            //strSql.AppendLine("      , B.BANK_ACNT_NO                                                                                               ");
            //strSql.AppendLine("      , B.ACNT_HOLDER                                                                                                ");
            //strSql.AppendLine("   FROM SUGMF A                                                                                                      ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B                                                                                         ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD                                                                                        ");
            //strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1                                                                                         ");
            //strSql.AppendLine("     ON B.CHRG_ID = B1.EMP_ID                                                                                        ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD C                                                                                           ");
            //strSql.AppendLine("     ON B.BANK_CD = C.COM_CD                                                                                         ");
            //strSql.AppendLine("    AND C.CD_GB = 'BANK_CD'                                                                                          ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN '"+ dicParams["DATE_F"] + "' AND '"+ dicParams["DATE_T"] + "'                                                                ");
            //strSql.AppendLine("    AND(('"+ dicParams["AC_GB"] + "' = '0' AND 1 = 1)                                                                                        ");
            //strSql.AppendLine("         OR                                                                                                          ");
            //strSql.AppendLine("         ('"+ dicParams["AC_GB"] + "' = '1' AND A.JGUBN = '1')                                                                               ");
            //strSql.AppendLine("         OR                                                                                                          ");
            //strSql.AppendLine("         ('"+ dicParams["AC_GB"] + "' = '2' AND A.JGUBN = '2'))                                                                              ");
            //strSql.AppendLine("    AND(('"+ dicParams["FIND_WORD"] + "' = '' AND 1 = 1)                                                                                          ");
            //strSql.AppendLine("         OR                                                                                                          ");
            //strSql.AppendLine("         ('"+ dicParams["FIND_IDX"] + "' = '0' AND(B.DEALER_NM LIKE '%"+ dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%"+ dicParams["FIND_WORD"] + "%'))                                            ");
            //strSql.AppendLine("         OR                                                                                                          ");
            //strSql.AppendLine("         ('"+ dicParams["FIND_IDX"] + "' = '1' AND B.IDT_NO LIKE '%"+ dicParams["FIND_WORD"] + "%')                                                                          ");
            //strSql.AppendLine("         OR                                                                                                          ");
            //strSql.AppendLine("         ('"+ dicParams["FIND_IDX"] + "' = '2' AND B1.EMP_NM LIKE '%"+ dicParams["FIND_WORD"] + "%'))                                                                        ");
            //strSql.AppendLine("  ORDER BY A.SLIPNO                                                                                                  ");
            #endregion

            #region mariaDB
            //strSql.AppendLine(" SELECT A.SLIPNO ");
            //strSql.AppendLine("      , A.CVCOD ");
            //strSql.AppendLine("      , B.DEALER_NM ");
            //strSql.AppendLine("      , CASE WHEN A.JGUBN = '1' THEN '0251' ELSE '0253' END AS ACCOD ");
            //strSql.AppendLine("      , CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-', SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS IDT_NO ");
            //strSql.AppendLine("      , B1.EMP_NM AS CHRG_NM ");
            //strSql.AppendLine("      , A.TDATE ");
            //strSql.AppendLine("      , (SELECT X1.ATEXT FROM ACTRAN X1 WHERE X1.REF1 = A.SLIPNO AND X1.AAUTO = 'D01' LIMIT 1) AS RMK ");
            //strSql.AppendLine("      , 0 AS AMT ");
            ////strSql.AppendLine("      , FORMAT(A.TRSUM, 0) AS TRSUM ");
            //strSql.AppendLine("      , A.TRSUM AS TRSUM ");
            //strSql.AppendLine("      , 0 AS DIFF ");
            //strSql.AppendLine("      , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK ");
            //strSql.AppendLine("      , B.BANK_ACNT_NO  ");
            //strSql.AppendLine("      , B.ACNT_HOLDER  ");
            //strSql.AppendLine("   FROM SUGMF A ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            //strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1 ");
            //strSql.AppendLine("     ON B.CHRG_ID = B1.EMP_ID ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD C ");
            //strSql.AppendLine("     ON B.BANK_CD = C.COM_CD ");
            //strSql.AppendLine("    AND C.CD_GB = 'BANK_CD' ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T ");
            //strSql.AppendLine("    AND ((@AC_GB = '0' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '1' AND A.JGUBN = '1' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '2' AND A.JGUBN = '2' ))  ");
            //strSql.AppendLine("    AND ((@FIND_WORD = '' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '2' AND B1.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))  ");
            #endregion

            #region[2021-02-09 이전 쿼리]

            /*
            * 2021-01-13 현업요청
            * 기존방식 : 기준테이블을 SUGMF가 아닌 전표테이블 ACTRAN에서 바로 가져올수 있도록
            * 이유 : 외상대 지불을 지불관리에서 등록하지않고 전표관리에서 바로 KeyIn하는 경우가 있기때문에 기준을 전표테이블로 수정
            * 한계 : 외상대 지불을 구분하는 구분자가 없어 추후 전표관리 등록창 부분 수정해야할 것으로 예상
            */
            //strSql.Clear();
            //strSql.AppendLine(" SELECT A.CVCOD ");
            //strSql.AppendLine("      , B.DEALER_NM ");
            //strSql.AppendLine("      , A.ACCOD ");
            //strSql.AppendLine("      , CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-', SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS IDT_NO  ");
            //strSql.AppendLine("      , B.CHRG_NM ");
            //strSql.AppendLine("      , A.ATEXT AS RK ");
            //strSql.AppendLine("      , DATE_FORMAT(A.TDATE, '%Y-%m-%d') AS TDATE ");
            //strSql.AppendLine("      , 0 AS AMT ");
            //strSql.AppendLine("      , CASE WHEN A.ACCOD = '0251' THEN IF(A.ACAMT = 0 OR A.ACAMT IS NULL, A.ADAMT, A.ACAMT)  ELSE IF(A.ADAMT = 0 OR A.ADAMT IS NULL, A.ACAMT, A.ADAMT) END AS TRSUM ");
            //strSql.AppendLine("      , 0 AS DIFF ");
            //strSql.AppendLine("      , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK  ");
            //strSql.AppendLine("      , B.BANK_ACNT_NO   ");
            //strSql.AppendLine("      , B.ACNT_HOLDER ");
            //strSql.AppendLine("   FROM ACTRAN A ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B  ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            //strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1 ");
            //strSql.AppendLine("     ON B.CHRG_ID = B1.EMP_ID ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD C ");
            //strSql.AppendLine("     ON B.BANK_CD = C.COM_CD ");
            //strSql.AppendLine("    AND C.CD_GB = 'BANK_CD' ");
            //strSql.AppendLine("  WHERE (A.AAUTO = '' OR A.AAUTO = 'D01' OR A.AAUTO IS NULL) ");
            //strSql.AppendLine("    AND A.ATGUB = '3' ");
            //strSql.AppendLine("    AND (A.CVCOD <= 9000 OR CVCOD >= 10000) ");
            //strSql.AppendLine("    AND A.ACCOD IN ( '0251', '0253' ) ");
            //strSql.AppendLine("    AND ((@AC_GB = '0' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '1' AND A.ACCOD = '0251' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '2' AND A.ACCOD = '0253' ))  ");
            //strSql.AppendLine("    AND ((@FIND_WORD = '' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '0' AND (B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' OR B.INITIAL_NM LIKE '%" + dicParams["FIND_WORD"] + "%'))  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '2' AND B1.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))  ");
            //strSql.AppendLine("    AND A.TDATE BETWEEN DATE_FORMAT(@DATE_F, '%Y%m%d') AND DATE_FORMAT(@DATE_T, '%Y%m%d') ");

            #endregion[2021-02-09 이전 쿼리]

            #region[2021-01-13 이전 쿼리]

            //strSql.Clear();
            //strSql.AppendLine(" SELECT A.SLIPNO ");
            //strSql.AppendLine("      , A.CVCOD ");
            //strSql.AppendLine("      , B.DEALER_NM ");
            //strSql.AppendLine("      , CASE WHEN A.JGUBN = '1' THEN '0251' ELSE '0253' END AS ACCOD ");
            //strSql.AppendLine("      , CONCAT(SUBSTRING(B.IDT_NO, 1, 3), '-', SUBSTRING(B.IDT_NO, 4, 2), '-', SUBSTRING(B.IDT_NO, 6, 5)) AS IDT_NO ");
            //strSql.AppendLine("      , B1.EMP_NM AS CHRG_NM ");
            //strSql.AppendLine("      , A.TDATE ");
            //strSql.AppendLine("      , (SELECT X1.ATEXT FROM ACTRAN X1 WHERE X1.REF1 = A.SLIPNO AND X1.AAUTO = 'D01' LIMIT 1) AS RMK ");
            //strSql.AppendLine("      , 0 AS AMT ");
            //strSql.AppendLine("      , FORMAT(A.TRSUM, 0) AS TRSUM ");
            //strSql.AppendLine("      , 0 AS DIFF ");
            //strSql.AppendLine("      , CASE WHEN C.COM_NM IS NULL THEN B.BANK_CD ELSE C.COM_NM END AS BANK ");
            //strSql.AppendLine("      , B.BANK_ACNT_NO  ");
            //strSql.AppendLine("      , B.ACNT_HOLDER  ");
            //strSql.AppendLine("   FROM SUGMF A ");
            //strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            //strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B1 ");
            //strSql.AppendLine("     ON B.CHRG_ID = B1.EMP_ID ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD C ");
            //strSql.AppendLine("     ON B.BANK_CD = C.COM_CD ");
            //strSql.AppendLine("    AND C.CD_GB = 'BANK_CD' ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN @DATE_F AND @DATE_T ");
            //strSql.AppendLine("    AND ((@AC_GB = '0' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '1' AND A.JGUBN = '1' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@AC_GB = '2' AND A.JGUBN = '2' ))  ");


            //strSql.AppendLine("    AND ((@FIND_WORD = '' AND 1 = 1)  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '0' AND B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '1' AND B.IDT_NO LIKE '%" + dicParams["FIND_WORD"] + "%' )  ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         (@FIND_IDX = '2' AND B1.EMP_NM LIKE '%" + dicParams["FIND_WORD"] + "%' ))  ");

            #endregion[2021-01-13 이전 쿼리]

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 프린트 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (GridViewRetr.RowCount > 0)
                ComnEtcFunc.ExportExcelFile(string.Format("{0}_", this.Text), GridRetr);
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 프린트 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("USRCD", FmMainToolBar2.UserID);

            ArrayList arr = new ArrayList();
            
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT B.EMP_NM ");
            strSql.AppendLine("   FROM ZUSRLST A ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B  ");
            strSql.AppendLine("     ON A.INSANO = B.EMP_ID ");
            strSql.AppendLine("  WHERE A.USRCD = @USRCD ");

            DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
            if(dt.Rows.Count > 0)
            {
                arr.Add(dt.Rows[0]["EMP_NM"]?.ToString());
            }
            else
            {
                arr.Add("");
            }

            arr.Add(DateTime.Now.ToString());
            arr.Add(string.Format("{0} ~ {1}", DateEditFrom.EditValue?.ToString().Substring(0, 10), DateEditTo.EditValue?.ToString().Substring(0, 10)));

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            int iFindIdx = CboFindSbj.SelectedIndex;
            string sFindWord = TxtFindWord.EditValue?.ToString();
            if (string.IsNullOrEmpty(TxtFindWord.EditValue?.ToString()))
                sFindWord = "";
            int iJGubn = CboAcNam.SelectedIndex;

            dicParams.Clear();
            dicParams.Add("DATE_F", sYmdFrom);
            dicParams.Add("DATE_T", sYmdTo);
            dicParams.Add("FIND_IDX", iFindIdx.ToString());
            dicParams.Add("FIND_WORD", sFindWord);
            dicParams.Add("AC_GB", iJGubn.ToString());

            dt = GetInfo(dicParams, "");

            if(dt.Rows.Count == 0)
            {
                string sMsg = string.Format("조회기간 : {0}" +
                    "\r\n해당기간 동안 발생된 지출내역이 존재하지 않습니다.", arr[1].ToString());

                XtraMessageBox.Show(sMsg);
                return;
            }

            int iTotAmt = 0;
            foreach(DataRow row in dt.Rows)
            {
                string sVal = row["TRSUM"]?.ToString();
                int i = string.IsNullOrEmpty(sVal.Replace(",", "")) ? 0 : Convert.ToInt32(sVal.Replace(",", ""));
                iTotAmt += i;
            }

            #region 2022-10-11 이후 주석 #0003
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    string sCvCod = dt.Rows[i]["CVCOD"]?.ToString();
            //    string sAcCod = dt.Rows[i]["ACCOD"]?.ToString();
            //    string sTrSum = dt.Rows[i]["TRSUM"]?.ToString();

            //    double dTrSum = string.IsNullOrEmpty(sTrSum) ? 0 : Convert.ToDouble(sTrSum);

            //    dicParams.Clear();
            //    dicParams.Add("DATE_F", DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8));
            //    dicParams.Add("DATE_T", DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8));
            //    dicParams.Add("ACCOD", sAcCod);
            //    dicParams.Add("CVCOD", sCvCod);

            //    DataTable dtTemp = GetInfo2(dicParams);
            //    if(dtTemp.Rows.Count > 0)
            //    {
            //        string sOfsBlc = dtTemp.Rows[0]["OFS_BLC"]?.ToString();
            //        double dOfsBls = string.IsNullOrEmpty(sOfsBlc) ? 0 : Convert.ToDouble(sOfsBlc);

            //        //#0003
            //        double dAmt = dOfsBls + dTrSum;
            //        dt.Rows[i]["AMT"] = dAmt;
            //        dt.Rows[i]["DIFF"] = dAmt - dTrSum;

            //        //#0002
            //        //double dAmt = 0;
            //        //if (dOfsBls != 0)
            //        //{
            //        //    dAmt = dOfsBls - dTrSum;
            //        //}
            //        //dt.Rows[i]["AMT"] = dOfsBls;
            //        //dt.Rows[i]["DIFF"] = dAmt;
            //    }
            //    else
            //    {

            //    }
            //}
            #endregion
            //DataTable copy = dt.Clone();

            //copy.Columns["AMT"].DataType = typeof(string);
            //foreach(DataRow row in dt.Rows)
            //{
            //    row["AMT"] = string.Format("{0:n0}", row["AMT"]?.ToString());
            //    copy.ImportRow(row);
            //}

            arr.Add(string.Format("{0:n0}", iTotAmt));

            ReportViewer fm = new ReportViewer(dt, arr, "RptVoucher");
            fm.ShowDialog();
        }

        /// <summary>
        ///     거래처별 미지급 현황 조회
        /// </summary>
        /// <param name="dicParams"> 쿼리변수 BtnRetr_Click 참조 </param> 
        private DataTable GetInfo2(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
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
            //strSql.AppendLine("                     AND A1.CVCOD = @CVCOD ");
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
            ////;strSql.AppendLine("                     AND A1.APVYN = 'Y' ");
            //strSql.AppendLine("                     AND A1.CVCOD = @CVCOD ");
            //strSql.AppendLine("                   GROUP BY A1.CVCOD ");
            //strSql.AppendLine("                ) X1 ");
            //strSql.AppendLine("          GROUP BY X1.CVCOD ");
            //strSql.AppendLine("    )    ");
            //strSql.AppendLine(" SELECT Y1.CVCOD  ");
            //strSql.AppendLine("      , Y1.JAMT + Y1.OCR_AMT - Y1.CPT_OFS_AMT AS OFS_BLC  ");
            //strSql.AppendLine("   FROM (  ");
            //strSql.AppendLine("          SELECT X1.CVCOD  ");
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
            ////strSql.AppendLine("             AND X1.APVYN = 'Y'  ");
            //strSql.AppendLine("             AND X1.CVCOD = @CVCOD    ");
            //strSql.AppendLine("           GROUP BY X1.CVCOD ");
            //strSql.AppendLine("        ) Y1 ");
            //strSql.AppendLine("  WHERE Y1.CVCOD = @CVCOD ");
            #endregion
                                                                                                                                                                                                
            strSql.AppendLine("WITH INFO AS(                         ");
            strSql.AppendLine("         SELECT X1.CVCOD              ");
            strSql.AppendLine("              , X1.ACCOD              ");
            strSql.AppendLine("              , SUM(X1.JAMT) AS JAMT  ");
            strSql.AppendLine("           FROM(                      ");
            strSql.AppendLine("                  SELECT A1.ACCOD     ");
            strSql.AppendLine("                       , A1.CVCOD     ");
            strSql.AppendLine("                       , ISNULL(MAX(A1.ACDRJN), 0) + ISNULL(MAX(A1.ACCRJN), 0) AS JAMT ");
            strSql.AppendLine("                    FROM ACJANF A1                                                     ");
            strSql.AppendLine("                   WHERE A1.ACYEAR = SUBSTRING('" + dicParams["DATE_F"] + "',1,4)   ");
            strSql.AppendLine("                     AND A1.ACCOD = '" + dicParams["ACCOD"] + "'  ");
            strSql.AppendLine("                     AND A1.CVCOD = " + dicParams["CVCOD"] + " ");
            strSql.AppendLine("                   GROUP BY A1.CVCOD, A1.ACCOD                                                                                                                   ");
            strSql.AppendLine("                   UNION ALL                                                                                                                                     ");
            strSql.AppendLine("                  SELECT A1.ACCOD                                                                                                                                ");
            strSql.AppendLine("                       , A1.CVCOD                                                                                                                                ");
            strSql.AppendLine("                       , ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ISNULL(ACAMT, 0) - ISNULL(ADAMT, 0) ELSE ISNULL(ADAMT, 0) - ISNULL(ACAMT, 0) END), 0) AS JAMT  ");
            strSql.AppendLine("                    FROM ACTRAN A1                                                                                                                               ");
            strSql.AppendLine("                    LEFT JOIN ACMSTF B1                                                                                                                          ");
            strSql.AppendLine("                      ON A1.ACCOD = B1.ACCOD                                                                                                                     ");
            strSql.AppendLine("                   WHERE A1.TDATE BETWEEN SUBSTRING('" + dicParams["DATE_F"] + "',1,4) + '0101' AND CONVERT(VARCHAR(8),DATEADD(DAY,-1,CONVERT(DATE,'" + dicParams["DATE_F"] + "')), 112)  ");
            strSql.AppendLine("                     AND A1.ACCOD = '" + dicParams["ACCOD"] + "'    ");
            //strSql.AppendLine("                     AND A1.APVYN = 'Y' ");//#0001 //#0003
            strSql.AppendLine("                     AND A1.CVCOD = " + dicParams["CVCOD"]);
            strSql.AppendLine("                   GROUP BY A1.CVCOD, A1.ACCOD               ");
            strSql.AppendLine("                ) X1                                         ");
            strSql.AppendLine("          GROUP BY X1.CVCOD, X1.ACCOD                        ");
            strSql.AppendLine("    )                                                        ");
            strSql.AppendLine("    SELECT Y1.CVCOD                                        ");
            strSql.AppendLine("         , Y2.JAMT + Y1.OCR_AMT - Y1.CPT_OFS_AMT AS OFS_BLC");
            strSql.AppendLine("     FROM(                                                 ");
            strSql.AppendLine("             SELECT X1.CVCOD                               ");
            strSql.AppendLine("                  , SUM(ISNULL(X1.ADAMT, 0)) AS OCR_AMT    ");
            strSql.AppendLine("                  , SUM(ISNULL(X1.ACAMT, 0)) AS CPT_OFS_AMT");
            strSql.AppendLine("                  , 0 AS OFS_AMT                           ");
            strSql.AppendLine("               FROM ACTRAN X1                              ");
            strSql.AppendLine("               LEFT JOIN ACC_DEALER_CD X2                  ");
            strSql.AppendLine("                 ON X1.CVCOD = X2.DEALER_CD                ");
            strSql.AppendLine("              WHERE X1.TDATE BETWEEN '" + dicParams["DATE_F"] + "' AND '" + dicParams["DATE_T"] + "'");
            strSql.AppendLine("                AND X1.ACCOD = '" + dicParams["ACCOD"] + "' --HARDCODING          ");
            //strSql.AppendLine("                AND X1.APVYN = 'Y'                         "); //#0003
            strSql.AppendLine("                AND X1.CVCOD = " + dicParams["CVCOD"]);
            strSql.AppendLine("              GROUP BY X1.CVCOD                            ");
            strSql.AppendLine("           ) Y1                                            ");
            strSql.AppendLine("     LEFT JOIN INFO Y2                                     ");
            strSql.AppendLine("       ON Y1.CVCOD = Y2.CVCOD                              ");
            strSql.AppendLine("    WHERE Y1.CVCOD = " + dicParams["CVCOD"]);

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC17001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                return;
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
            else if (e.KeyCode == Keys.F9)
                BtnPrint.PerformClick();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AC17001F01_TextChanged(object sender, EventArgs e)
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

        private void CboAcNam_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBoxEdit cbo = (ComboBoxEdit)sender;
            if (cbo.SelectedIndex < 0)
                return;

            BtnRetr.PerformClick();
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}