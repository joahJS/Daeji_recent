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
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraBars;
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
*            
* 수정일자 : 2023-02-10
* 수정자   : 정은영
* ID       : #0001
* 수정내용 : (현업요청)
*            1. 잔액이 있는 계정 모두 표시
*/
namespace AccAdm
{
    public partial class AC08001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC08001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC08001F01_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            DateEditFrom.EditValue = today.AddDays(1 - today.Day);
            DateEditTo.EditValue = today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewAcc, GridViewCv };
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


            InitControls();
            UpdateDropDownButton(BtnSlip);
            BtnRetr_Click(null, null);
        }

        private void GridViewAcc_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (GridViewAcc.RowCount == 0)
                return;

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sAcCod = GridViewAcc.GetFocusedRowCellValue("ACCOD")?.ToString();

            GridCv.DataSource = GetAccountDetail(sYmdFrom, sYmdTo, sAcCod);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sAcCod = BtnEditAcCod.EditValue?.ToString().Trim();

            /*
             * 2021-01-06 현업요청
             * 원장의 경우 잔액부분 때문에 조회일자 From이 2021년 이전일 경우 메시지 출력 및 2021년도로 다시 세팅
             */
            DateTime dateFrom = DateTime.Parse(DateEditFrom.EditValue?.ToString());
            DateTime dateTo = DateTime.Parse(DateEditTo.EditValue?.ToString());
            if(dateFrom.Year < 2021 || dateTo.Year < 2021)
            {
                XtraMessageBox.Show("계정별 원장은 2021년 01월 01일부터 조회가 가능합니다.");
                DateEditFrom.EditValue = dateFrom.AddYears(2021 - dateFrom.Year);
                DateEditTo.EditValue = dateTo.AddYears(2021 - dateTo.Year);
                return;
            }

            GridAcc.DataSource = null;
            GridAcc.DataSource = GetAccountInfo(sYmdFrom, sYmdTo, sAcCod);
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        #region[Execute By Query]

        private DataTable GetAccountInfo(string sYmdFrom, string sYmdTo, string sAcCod)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            //#0001
            strSql.AppendLine("SELECT REPLACE(Z1.ACCOD,' ','') AS ACCOD                                                                                                                             ");
            strSql.AppendLine("     , Z2.ACDSP                                                                                                                                                      ");
            strSql.AppendLine("     , SUM(Z1.JAMT) AS JAMT                                                                                                                                          ");
            strSql.AppendLine(" FROM(SELECT A1.ACCOD                                                                                                                                                ");
            strSql.AppendLine("             , ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT                                                                                                      ");
            strSql.AppendLine("         FROM ACJANF A1                                                                                                                                              ");
            strSql.AppendLine("         WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'                                                                                                                                    ");
            strSql.AppendLine("         UNION ALL                                                                                                                                                   ");
            strSql.AppendLine("         SELECT A1.ACCOD                                                                                                                                             ");
            strSql.AppendLine("             , CASE WHEN B1.ACRDR = '1' THEN SUM(ISNULL(A1.ACAMT, 0)) - SUM(ISNULL(A1.ADAMT, 0)) ELSE SUM(ISNULL(A1.ADAMT, 0)) - SUM(ISNULL(A1.ACAMT, 0)) END JAMT   ");
            strSql.AppendLine("         FROM ACTRAN A1                                                                                                                                              ");
            strSql.AppendLine("         LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                                                                  ");
            strSql.AppendLine("         WHERE A1.TDATE BETWEEN SUBSTRING('" + sYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + sYmdFrom + "')), 112)                       ");
            strSql.AppendLine("         GROUP BY A1.ACCOD, B1.ACRDR                                                                                                                                 ");
            strSql.AppendLine("         UNION ALL                                                                                                                                                   ");
            strSql.AppendLine("         --기간검색 전표                                                                                                                                             ");
            strSql.AppendLine("         SELECT A.ACCOD                                                                                                                                              ");
            strSql.AppendLine("             , CASE WHEN B.ACRDR = '1' THEN SUM(ISNULL(A.ACAMT, 0)) - SUM(ISNULL(A.ADAMT, 0)) ELSE SUM(ISNULL(A.ADAMT, 0)) - SUM(ISNULL(A.ACAMT, 0)) END JAMT        ");
            strSql.AppendLine("             FROM ACTRAN A                                                                                                                                           ");
            strSql.AppendLine("             LEFT OUTER JOIN ACMSTF B                                                                                                                                ");
            strSql.AppendLine("             ON A.ACCOD = B.ACCOD                                                                                                                                    ");
            strSql.AppendLine("         WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'                                                                                                             ");
            strSql.AppendLine("             AND A.ACCOD <> '" + ComnEtcFunc.CashCode + "'                                                                                                                                   ");
            strSql.AppendLine("             AND A.APVYN = 'Y'                                                                                                                                       ");
            strSql.AppendLine("         GROUP BY A.ACCOD, B.ACRDR)Z1                                                                                                                                ");
            strSql.AppendLine(" LEFT JOIN ACMSTF Z2                                                                                                                                                 ");
            strSql.AppendLine(" ON Z1.ACCOD = Z2.ACCOD                                                                                                                                              ");
            strSql.AppendLine(" WHERE Z1.ACCOD <> '" + ComnEtcFunc.CashCode + "'                                                                                                                                            ");
            strSql.AppendLine("   AND (('" + sAcCod + "' = '' AND 1 = 1 ) ");
            strSql.AppendLine("        OR ");
            strSql.AppendLine("        ('" + sAcCod + "' <> '' AND Z1.ACCOD = '" + sAcCod + "' ))");
            strSql.AppendLine(" GROUP BY Z1.ACCOD, Z2.ACDSP                                                                                                                                         ");
            strSql.AppendLine(" HAVING SUM(Z1.JAMT) <> 0                                                                                                                                            ");
            strSql.AppendLine(" ORDER BY Z1.ACCOD                                                                                                                                                   ");

            #region #0001 이전
            //strSql.AppendLine(" SELECT DISTINCT A.ACCOD ");
            //strSql.AppendLine(" 	 , B.ACDSP ");
            //strSql.AppendLine("   FROM ACTRAN A ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            //strSql.AppendLine("     ON A.ACCOD = B.ACCOD ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.ACCOD <> '" + ComnEtcFunc.CashCode + "' ");
            //strSql.AppendLine("    AND (('" + sAcCod + "' = '' AND 1 = 1 ) ");
            //strSql.AppendLine("         OR ");
            //strSql.AppendLine("         ('" + sAcCod + "' <> '' AND A.ACCOD = '" + sAcCod + "' ))");
            //strSql.AppendLine("    AND A.APVYN = 'Y' ");
            //strSql.AppendLine("  ORDER BY A.ACCOD ");
            #endregion

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetAccountDetail(string sYmdFrom, string sYmdTo, string sAcCod)
        {
            StringBuilder strSql = new StringBuilder();

            #region mariaDB
            //strSql.AppendLine(" # 계정별 원장 ");
            //strSql.AppendLine(" SELECT A1.TDATE AS TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, CVCOD, CVNAM, ACAMT, ADAMT ");
            //strSql.AppendLine("      , SUM((CASE WHEN ACRDR='1' THEN ACAMT-ADAMT ELSE ADAMT-ACAMT END)+JAMT) OVER(ORDER BY TDATE, SEQNO, LINNO) JJAMT   ");
            //strSql.AppendLine("   FROM ( SELECT ' [이월]' TDATE, ' 'SEQNO, 0 LINNO, '' ACRDR, ' 'ATEXT, '' CVCOD, '' CVNAM, 0 ACAMT,0 ADAMT, SUM(JAMT)JAMT ");
            //strSql.AppendLine("            FROM ( ");
            //strSql.AppendLine("                   SELECT IFNULL(A1.ACDRJN,0)+IFNULL(A1.ACCRJN,0) JAMT ");
            //strSql.AppendLine("                     FROM ACJANF A1  ");
            //strSql.AppendLine("                    WHERE A1.ACYEAR='" + sYmdFrom.Substring(0, 4) + "' AND A1.ACCOD='" + sAcCod + "'           ");
            //strSql.AppendLine("                    UNION ALL ");
            //strSql.AppendLine("                   SELECT IFNULL(SUM(CASE WHEN B1.ACRDR='1' THEN ACAMT-ADAMT ELSE ADAMT-ACAMT END),0)JAMT ");
            //strSql.AppendLine("                     FROM ACTRAN A1 ");
            //strSql.AppendLine("                     LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD ");
            //strSql.AppendLine("                    WHERE A1.TDATE BETWEEN DATE_FORMAT('" + sYmdFrom + "' ,'%Y0101') AND DATE_FORMAT(DATE_ADD(DATE_FORMAT('" + sYmdFrom + "', '%Y%m%d'), INTERVAL -1 DAY), '%Y%m%d') ");
            //strSql.AppendLine("                      AND A1.ACCOD='" + sAcCod + "'      ");
            //strSql.AppendLine("                 ) A1         ");
            //strSql.AppendLine("           UNION  ALL ");
            //strSql.AppendLine("          SELECT DATE_FORMAT(A1.TDATE, '%Y-%m-%d'), A1.SEQNO, A1.LINNO, B1.ACRDR, A1.ATEXT, A1.CVCOD, A1.CVNAM, IFNULL(A1.ACAMT, 0) AS ACAMT ,IFNULL(A1.ADAMT, 0) AS ADAMT, 0 JAMT ");
            //strSql.AppendLine("            FROM ACTRAN A1 ");
            //strSql.AppendLine("            LEFT JOIN ACMSTF B1 ON A1.ACCOD=B1.ACCOD ");
            //strSql.AppendLine("           WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("             AND A1.ACCOD='" + sAcCod + "' ");
            //strSql.AppendLine("         ) A1       ");
            #endregion

            strSql.AppendLine("SELECT A1.TDATE AS TDATE, A1.SEQNO, A1.LINNO, A1.ATEXT, CVCOD, CVNAM, ACAMT, ADAMT                                                                                                                 ");
            strSql.AppendLine("     , SUM((CASE WHEN ACRDR = '1' THEN ACAMT - ADAMT ELSE ADAMT - ACAMT END) + JAMT) OVER(ORDER BY TDATE, SEQNO, LINNO) JJAMT                                                                      ");
            strSql.AppendLine("  FROM(SELECT ' [이월]' TDATE, ' 'SEQNO, 0 LINNO, '' ACRDR, ' 'ATEXT, '' CVCOD, '' CVNAM, 0 ACAMT, 0 ADAMT, SUM(JAMT)JAMT                                                                          ");
            strSql.AppendLine("           FROM(                                                                                                                                                                                   ");
            strSql.AppendLine("                  SELECT ISNULL(A1.ACDRJN, 0) + ISNULL(A1.ACCRJN, 0) JAMT                                                                                                                          ");
            strSql.AppendLine("                    FROM ACJANF A1                                                                                                                                                                 ");
            strSql.AppendLine("                   WHERE A1.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "' AND A1.ACCOD = '" + sAcCod + "'                                                                                                                                  ");
            strSql.AppendLine("                   UNION ALL                                                                                                                                                                       ");
            //strSql.AppendLine("                  SELECT ISNULL(SUM(CASE WHEN B1.ACRDR = '1' THEN ACAMT - ADAMT ELSE ADAMT - ACAMT END), 0)JAMT                                                                                    ");
            strSql.AppendLine("                  SELECT CASE WHEN B1.ACRDR = '1' THEN SUM(ISNULL(A1.ACAMT,0)) -SUM(ISNULL(A1.ADAMT, 0)) ELSE SUM(ISNULL(A1.ADAMT,0)) -SUM(ISNULL(A1.ACAMT, 0)) END JAMT ");
            strSql.AppendLine("                    FROM ACTRAN A1                                                                                                                                                                 ");
            strSql.AppendLine("                    LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                                                                                     ");
            strSql.AppendLine("                   WHERE A1.TDATE BETWEEN SUBSTRING('" + sYmdFrom + "', 1, 4) + '0101' AND CONVERT(VARCHAR(8), DATEADD(DAY, -1, CONVERT(DATE, '" + sYmdFrom + "')), 112)                                           ");
            strSql.AppendLine("                     AND A1.ACCOD = '" + sAcCod + "'                                                                                                                                                         ");
            strSql.AppendLine("                   GROUP BY B1.ACRDR");
            strSql.AppendLine("                ) A1                                                                                                                                                                               ");
            strSql.AppendLine("          UNION  ALL                                                                                                                                                                               ");
            strSql.AppendLine("         SELECT CONVERT(VARCHAR(10), CONVERT(DATE, A1.TDATE), 23), A1.SEQNO, A1.LINNO, B1.ACRDR, A1.ATEXT, A1.CVCOD, A1.CVNAM, ISNULL(A1.ACAMT, 0) AS ACAMT, ISNULL(A1.ADAMT, 0) AS ADAMT, 0 JAMT  ");
            strSql.AppendLine("           FROM ACTRAN A1                                                                                                                                                                          ");
            strSql.AppendLine("           LEFT JOIN ACMSTF B1 ON A1.ACCOD = B1.ACCOD                                                                                                                                              ");
            strSql.AppendLine("          WHERE A1.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'                                                                                                                                         ");
            strSql.AppendLine("            AND A1.APVYN = 'Y'");
            strSql.AppendLine("            AND A1.ACCOD = '" + sAcCod + "'                                                                                                                                                                  ");
            strSql.AppendLine("        ) A1                                                                                                                                                                                       ");
            strSql.AppendLine("  ORDER BY A1.TDATE, A1.SEQNO, A1.LINNO                                                                                                                                                            ");
                                                                                                                                                                                                                                  
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());                                                                                                                                                          

            #region 이전코드
            //strSql.Clear();
            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT '' AS TDATE ");
            //strSql.AppendLine("      , '전월 이월' AS ATEXT  ");
            //strSql.AppendLine("      , NULL AS CVCOD  ");
            //strSql.AppendLine("      , '' AS CVNAM ");
            //strSql.AppendLine("      , NULL AS ACAMT ");
            //strSql.AppendLine("      , NULL AS ADAMT ");
            //strSql.AppendLine("      , IFNULL(ACJANF_AMT, 0) - IFNULL(ACTRAN_AMT, 0) AS BAL_AMT ");
            //strSql.AppendLine("  FROM( ");
            //strSql.AppendLine("        SELECT CASE WHEN ROWNUM = '1' THEN IFNULL(CARRY_OVER_AMT1, 0) END AS ACJANF_AMT ");
            //strSql.AppendLine("        	 , CASE WHEN ROWNUM = '2' THEN IFNULL(CARRY_OVER_AMT2, 0) END AS ACTRAN_AMT ");
            //strSql.AppendLine("          FROM ( ");
            //strSql.AppendLine("                 SELECT '1' AS ROWNUM ");
            //strSql.AppendLine("                      , IFNULL(SUM(IFNULL(A.ACDRJN, 0) + IFNULL(A.ACCRJN, 0)), 0) AS CARRY_OVER_AMT1 ");
            //strSql.AppendLine("                 	  , 0 AS CARRY_OVER_AMT2 ");
            //strSql.AppendLine("                   FROM ACJANF A  ");
            //strSql.AppendLine("                   LEFT OUTER JOIN ACMSTF B ");
            //strSql.AppendLine("                     ON A.ACCOD = B.ACCOD ");
            //strSql.AppendLine("                  WHERE A.ACYEAR = '" + sYmdFrom.Substring(0, 4) + "'    ");
            //strSql.AppendLine("                    AND A.ACCOD = '" + sAcCod + "' ");
            //strSql.AppendLine("                  GROUP BY A.ACCOD ");
            //strSql.AppendLine("                  UNION ALL ");
            //strSql.AppendLine("                 SELECT '2' AS ROWNUM ");
            //strSql.AppendLine("                      , 0 AS CARRY_OVER_AMT1 ");
            //strSql.AppendLine("                      , IFNULL(SUM(IFNULL(C.ADAMT, 0) - IFNULL(C.ACAMT, 0)), 0) AS CARRY_OVER_AMT2 ");
            //strSql.AppendLine("                   FROM ACTRAN C ");
            //strSql.AppendLine("                   LEFT OUTER JOIN ACMSTF D");
            //strSql.AppendLine("                     ON C.ACCOD = D.ACCOD");
            //strSql.AppendLine("                  WHERE C.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("                    AND C.ACCOD = '" + sAcCod + "'    ");
            //strSql.AppendLine("                  GROUP BY D.ASMCD ");
            //strSql.AppendLine("               ) X1    ");
            //strSql.AppendLine("      ) X2  ");

            //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //strSql.Clear();
            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT A.TDATE  ");
            //strSql.AppendLine(" 	 , A.ATEXT ");
            //strSql.AppendLine(" 	 , A.CVCOD ");
            //strSql.AppendLine(" 	 , B.DEALER_NM AS CVNAM ");
            //strSql.AppendLine(" 	 , A.ACAMT ");
            //strSql.AppendLine(" 	 , A.ADAMT ");
            //strSql.AppendLine(" 	 , 0 AS BAL_AMT ");
            //strSql.AppendLine("   FROM ACTRAN A  ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B ");
            //strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD  ");
            //strSql.AppendLine("  WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'  ");
            //strSql.AppendLine("    AND A.ACCOD = '" + sAcCod + "'  ");
            //strSql.AppendLine("  ORDER BY A.ACCOD  ");

            //DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //DataTable dtCopy = dt2.Clone();
            //dtCopy.TableName = "Table1";
            //dtCopy.ImportRow(dt.Rows[0]);
            //for (int i = 0; i < dt2.Rows.Count; i++)
            //{
            //    dtCopy.ImportRow(dt2.Rows[i]);
            //    if (i == 0)
            //        continue;

            //    string sPreBalAmt = string.IsNullOrEmpty(dtCopy.Rows[i - 1]["BAL_AMT"].ToString()) ? "0" : dtCopy.Rows[i - 1]["BAL_AMT"].ToString();
            //    string sCurrDeposit = string.IsNullOrEmpty(dtCopy.Rows[i]["ADAMT"].ToString()) ? "0" : dtCopy.Rows[i]["ADAMT"].ToString();
            //    string sCurrWithDraw = string.IsNullOrEmpty(dtCopy.Rows[i]["ACAMT"].ToString()) ? "0" : dtCopy.Rows[i]["ACAMT"].ToString();

            //    double dPreBalAmt = Convert.ToDouble(sPreBalAmt);
            //    double dCurrDeposit = Convert.ToDouble(sCurrDeposit);
            //    double dCurrWithDraw = Convert.ToDouble(sCurrWithDraw);

            //    dtCopy.Rows[i]["BAL_AMT"] = dPreBalAmt + (dCurrDeposit - dCurrWithDraw);
            //}

            //return dtCopy;
            #endregion
        }

        #endregion[Execute By Query]

        #region[GridView's Row Design]
        private void GridViewAcc_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewAcc_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        #endregion[GridView's Row Design]

        private void AC08001F01_KeyDown(object sender, KeyEventArgs e)
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
                DropBtnExcel.PerformClick();
            }
        }

        private void AC08001F01_TextChanged(object sender, EventArgs e)
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

        public DataRow DrPopupInfo;
        private void BtnEditAcCod_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                ButtonEdit btnEdit = (ButtonEdit)sender;
                string sVal = btnEdit.EditValue?.ToString().Trim();

                if (string.IsNullOrEmpty(sVal))
                {
                    TxtAcNam.EditValue = string.Empty;
                    return;
                }

                DataTable dt = GetAccInfo(sVal);
                if(dt.Rows.Count == 1)
                {
                    btnEdit.EditValue = dt.Rows[0]["ACCOD"];
                    TxtAcNam.EditValue = dt.Rows[0]["ACDSP"];
                }
                else
                {
                    AC01001F03 frm = new AC01001F03();
                    frm.P_AC08001F01 = this;
                    frm.AccCd = sVal;
                    if(frm.ShowDialog() == DialogResult.OK)
                    {
                        btnEdit.EditValue = DrPopupInfo["ACCOD"];
                        TxtAcNam.EditValue = DrPopupInfo["ACDSP"];
                    }
                }

            }
        }

        private void BtnEditAcCod_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString().Trim();

            AC01001F03 frm = new AC01001F03();
            frm.P_AC08001F01 = this;
            frm.AccCd = sVal;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                btnEdit.EditValue = DrPopupInfo["ACCOD"];
                TxtAcNam.EditValue = DrPopupInfo["ACDSP"];
            }
        }

        private DataTable GetAccInfo(string sVal)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.ACCOD, A.ACNAM, A.ACDSP");
            strSql.AppendLine("   FROM ACMSTF A");
            strSql.AppendLine("  WHERE A.ACCOD = '" + sVal + "'");
            strSql.AppendLine("     OR A.ACNAM LIKE '%" + sVal + "%'");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewAcc_MouseDown(object sender, MouseEventArgs e)
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
        
        BarManager barManager1;
        PopupMenu popupMenu1;
        BarButtonItem BtnAcCod;
        BarButtonItem BtnSlip;

        private void InitControls()
        {
            barManager1 = new BarManager();
            barManager1.Form = this;
            
            popupMenu1 = new PopupMenu(barManager1);
            BtnAcCod = new BarButtonItem(barManager1, "계정코드");
            BtnSlip = new BarButtonItem(barManager1, "계정별원장");
            popupMenu1.AddItem(BtnAcCod);
            popupMenu1.AddItem(BtnSlip);

            DropBtnExcel.DropDownControl = popupMenu1;
            
            BtnAcCod.Tag = "계정코드";
            BtnAcCod.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnAcCod_ItemClick);
            // 
            // btnZoomOut
            // 
            BtnSlip.Tag = "계정별원장";
            BtnSlip.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BtnSlip_ItemClick);
        }

        private void BtnAcCod_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            ExcelAcCod();
        }

        private void BtnSlip_ItemClick(object sender, ItemClickEventArgs e)
        {
            UpdateDropDownButton(e.Item);
            //...
            ExcelSlip();
        }

        private void UpdateDropDownButton(BarItem submenuItem)
        {
            DropBtnExcel.ImageOptions.SvgImage = submenuItem.ImageOptions.SvgImage;
            DropBtnExcel.ImageOptions.SvgImageSize = new Size(16, 16);
            DropBtnExcel.Tag = submenuItem.Tag;
        }
        
        private void ExcelAcCod()
        {
            ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}", this.Text, "_코드리스트"), GridAcc);
        }

        private void ExcelSlip()
        {
            ComnEtcFunc.ExportExcelFile(string.Format("{0}_{1}", this.Text, "_원장리스트"), GridCv);
        }

        private void DropBtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string tag = (sender as DropDownButton).Tag?.ToString();
            if (tag == "계정코드")
            {
                ExcelAcCod();
            }
            else if (tag == "계정별원장")
            {
                ExcelSlip();
            }
        }
    }
}