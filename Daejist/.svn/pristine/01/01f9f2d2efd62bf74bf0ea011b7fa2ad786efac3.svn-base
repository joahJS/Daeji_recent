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
using System.Data.SqlClient;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * --------------------HISTORY------------------
 * 1. 2021-02-07 (현업요청)
 *    거래처 초성검색 추가
 * 
 * 
 */
namespace AccAdm
{
    public partial class AC18001F02 : DevExpress.XtraEditors.XtraForm
    {
        public AC18001F02()
        {
            InitializeComponent();
        }

        public DataRow DR_INFO;
        private void AC18001F02_Load(object sender, EventArgs e)
        {
            SetLookup();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            
            if (DR_INFO == null)
            {
                DateEditTDate.EditValue = DateTime.Today;
                dicParams.Add("TAXNO", "");
            }
            else
            {
                dicParams.Add("TAXNO", DR_INFO["TAXNO"].ToString());

                string sAutGb = DR_INFO["AUTGB"]?.ToString();
                if (!string.IsNullOrEmpty(sAutGb) && sAutGb.Equals("1"))
                {
                    BtnSave.Enabled = false;

                    LayoutGroup.AppearanceGroup.ForeColor = Color.Red;
                    LayoutGroup.Text += "( 마감자료를 발행한 계산서 건은 수정이 불가합니다. 삭제 후 다시 발행하세요. )";

                    SetConrolsReadOnly();
                }
            }

            SetGridData(dicParams);

            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);
        }

        private void SetConrolsReadOnly()
        {
            LkupTaxGu.ReadOnly = true;
            LkupTGubn.ReadOnly = true;
            LkupPayGb.ReadOnly = true;
            BtnCvNam.ReadOnly = true;
            TxtCvCod.ReadOnly = true;
            DateEditTDate.ReadOnly = true;
            TxtRmk.ReadOnly = true;
            GridViewRetr.OptionsBehavior.Editable = false;
        }

        private void SetGridData(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.TAXNO ");
            strSql.AppendLine("      , A.TDATE ");
            strSql.AppendLine("      , A.TAXGU ");
            strSql.AppendLine("      , A.TGUBN ");
            strSql.AppendLine("      , A.PAYGB ");
            strSql.AppendLine("      , A.CVCOD ");
            strSql.AppendLine("      , B.DEALER_NM ");
            strSql.AppendLine("      , B.IDT_NO ");
            strSql.AppendLine("      , B.ADDR ");
            strSql.AppendLine("      , B.REP_NM ");
            strSql.AppendLine("      , B.BIZ_NM ");
            strSql.AppendLine("      , B.TYPE_NM ");
            strSql.AppendLine("      , ISNULL(TAMT1, 0) + ISNULL(TAMT2, 0) + ISNULL(TAMT3, 0) + ISNULL(TAMT4, 0) AS TAMT ");
            strSql.AppendLine("      , ISNULL(TVAT1, 0) + ISNULL(TVAT2, 0) + ISNULL(TVAT3, 0) + ISNULL(TVAT4, 0) AS TVAT ");
            strSql.AppendLine("      , A.RK ");
            strSql.AppendLine("   FROM TAXF A ");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B  ");
            strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendLine("  WHERE A.TAXNO = @TAXNO ");

            DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
            if(dt.Rows.Count > 0)
            {
                TxtTaxNo.EditValue = dt.Rows[0]["TAXNO"];
                DateEditTDate.EditValue = dt.Rows[0]["TDATE"];
                LkupTaxGu.EditValue = dt.Rows[0]["TAXGU"];
                LkupTGubn.EditValue = dt.Rows[0]["TGUBN"];
                LkupPayGb.EditValue = dt.Rows[0]["PAYGB"];
                TxtCvCod.EditValue = dt.Rows[0]["CVCOD"];
                BtnCvNam.EditValue = dt.Rows[0]["DEALER_NM"];
                TxtIdtNo.EditValue = dt.Rows[0]["IDT_NO"];
                TxtAddr.EditValue = dt.Rows[0]["ADDR"];
                TxtRepNm.EditValue = dt.Rows[0]["REP_NM"];
                TxtUpTae.EditValue = dt.Rows[0]["BIZ_NM"];
                TxtJongK.EditValue = dt.Rows[0]["TYPE_NM"];
                TxtTotAmt.EditValue = dt.Rows[0]["TAMT"];
                TxtTotVat.EditValue = dt.Rows[0]["TVAT"];
            }

            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine("  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD1, 2), '-', RIGHT(MMDD1, 2)) AS MMDD, ITNM1 AS ITNM, SPEC1 AS SPEC, TWGT1 AS TWGT ");
            //strSql.AppendLine("       , COST1 AS COST, TAMT1 AS TAMT, TVAT1 AS TVAT, TRMK1 AS TRMK ");
            //strSql.AppendLine("       , TAMT1 + TVAT1 AS TOT_AMT ");
            //strSql.AppendLine("    FROM TAXF X1 ");
            //strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            //strSql.AppendLine("   UNION ALL ");
            //strSql.AppendLine("  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD2, 2), '-', RIGHT(MMDD2, 2)), ITNM2, SPEC2, TWGT2 ");
            //strSql.AppendLine("       , COST2, TAMT2, TVAT2, TRMK2 ");
            //strSql.AppendLine("       , TAMT2 + TVAT2 AS TOT_AMT ");
            //strSql.AppendLine("    FROM TAXF X1 ");
            //strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            //strSql.AppendLine("   UNION ALL ");
            //strSql.AppendLine("  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD3, 2), '-', RIGHT(MMDD3, 2)), ITNM3, SPEC3, TWGT3 ");
            //strSql.AppendLine("       , COST3, TAMT3, TVAT3, TRMK3 ");
            //strSql.AppendLine("       , TAMT3 + TVAT3 AS TOT_AMT ");
            //strSql.AppendLine("    FROM TAXF X1 ");
            //strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            //strSql.AppendLine("   UNION ALL ");
            //strSql.AppendLine("  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD4, 2), '-', RIGHT(MMDD4, 2)), ITNM4, SPEC4, TWGT4 ");
            //strSql.AppendLine("       , COST4, TAMT4, TVAT4, TRMK4 ");
            //strSql.AppendLine("       , TAMT4 + TVAT2 AS TOT_AMT ");
            //strSql.AppendLine("    FROM TAXF X1 ");
            //strSql.AppendLine("   WHERE TAXNO = @TAXNO  ");
            #endregion

            strSql.AppendLine(" SELECT CASE WHEN ISNULL(MMDD1,'') != '' THEN SUBSTRING(MMDD1, 1, 2)+'-' + SUBSTRING(MMDD1, 3, 2) ELSE '' END AS MMDD, ITNM1 AS ITNM, SPEC1 AS SPEC, TWGT1 AS TWGT");
            strSql.AppendLine("      , COST1 AS COST, TAMT1 AS TAMT, TVAT1 AS TVAT, TRMK1 AS TRMK                                                                                                ");
            strSql.AppendLine("      , TAMT1 + TVAT1 AS TOT_AMT                                                                                                                                  ");
            strSql.AppendLine("   FROM TAXF X1                                                                                                                                                   ");
            strSql.AppendLine("  WHERE TAXNO = @TAXNO                                                                                                                                            ");
            strSql.AppendLine("  UNION ALL                                                                                                                                                       ");
            strSql.AppendLine(" SELECT CASE WHEN ISNULL(MMDD2, '') != '' THEN SUBSTRING(MMDD2, 1, 2)+'-' + SUBSTRING(MMDD2, 3, 2) ELSE '' END AS MMDD, ITNM2, SPEC2, TWGT2                       ");
            strSql.AppendLine("      , COST2, TAMT2, TVAT2, TRMK2                                                                                                                                ");
            strSql.AppendLine("      , TAMT2 + TVAT2 AS TOT_AMT                                                                                                                                  ");
            strSql.AppendLine("   FROM TAXF X1                                                                                                                                                   ");
            strSql.AppendLine("  WHERE TAXNO = @TAXNO                                                                                                                                            ");
            strSql.AppendLine("  UNION ALL                                                                                                                                                       ");
            strSql.AppendLine(" SELECT CASE WHEN ISNULL(MMDD3,'') != '' THEN SUBSTRING(MMDD3, 1, 2)+'-' + SUBSTRING(MMDD3, 3, 2) ELSE '' END , ITNM3, SPEC3, TWGT3                               ");
            strSql.AppendLine("      , COST3, TAMT3, TVAT3, TRMK3                                                                                                                                ");
            strSql.AppendLine("      , TAMT3 + TVAT3 AS TOT_AMT                                                                                                                                  ");
            strSql.AppendLine("   FROM TAXF X1                                                                                                                                                   ");
            strSql.AppendLine("  WHERE TAXNO = @TAXNO                                                                                                                                            ");
            strSql.AppendLine("  UNION ALL                                                                                                                                                       ");
            strSql.AppendLine(" SELECT CASE WHEN ISNULL(MMDD4,'') != '' THEN SUBSTRING(MMDD4, 1, 2)+'-' + SUBSTRING(MMDD4, 3, 2) ELSE '' END, ITNM4, SPEC4, TWGT4                                ");
            strSql.AppendLine("      , COST4, TAMT4, TVAT4, TRMK4                                                                                                                                ");
            strSql.AppendLine("      , TAMT4 + TVAT2 AS TOT_AMT                                                                                                                                  ");
            strSql.AppendLine("   FROM TAXF X1                                                                                                                                                   ");
            strSql.AppendLine("  WHERE TAXNO = @TAXNO                                                                                                                                            ");

            DataTable dt2 = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
            if(dt.Rows.Count == 0)
            {
                DataRow row = dt2.NewRow();
                dt2.Rows.Add(row);

                DataRow row1 = dt2.NewRow();
                dt2.Rows.Add(row1);

                DataRow row2 = dt2.NewRow();
                dt2.Rows.Add(row2);

                DataRow row3 = dt2.NewRow();
                dt2.Rows.Add(row3);
            }

            GridRetr.DataSource = dt2;

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {

            string sTaxNo = TxtTaxNo.EditValue?.ToString();
            string sTaxGu = LkupTaxGu.EditValue?.ToString();
            string sTGubn = LkupTGubn.EditValue?.ToString();
            string sPayGb = LkupPayGb.EditValue?.ToString();
            string sCvCOd = TxtCvCod.EditValue?.ToString();
            string sTDate = DateEditTDate.EditValue?.ToString().Substring(0, 10);
            string sRK = TxtRmk.EditValue?.ToString().Trim();

            if (string.IsNullOrEmpty(sTaxGu) || sTaxGu.Equals("****"))
            {
                XtraMessageBox.Show("계산서구분을 선택하세요.");
                LkupTaxGu.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sTGubn) || sTGubn.Equals("****"))
            {
                XtraMessageBox.Show("청구영수구분을 선택하세요.");
                LkupTGubn.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sPayGb) || sPayGb.Equals("****"))
            {
                XtraMessageBox.Show("지급구분을 선택하세요.");
                LkupPayGb.SelectAll();
                LkupPayGb.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sCvCOd))
            {
                XtraMessageBox.Show("거래처를 선택하세요.");
                BtnCvNam.SelectAll();
                BtnCvNam.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sTDate))
            {
                XtraMessageBox.Show("작성일자를 입력하세요.");
                DateEditTDate.SelectAll();
                DateEditTDate.Focus();
                return;
            }

            if (!ValidateGridViewEssentialData(GridViewRetr))
                return;

            DataTable dt = ReArrangeRowSequence((DataTable)GridRetr.DataSource);
            
            string sMMDD1 = dt.Rows[0]["MMDD"]?.ToString();
            sMMDD1 = string.IsNullOrEmpty(sMMDD1) ? null : sMMDD1.Replace("-", "");
            string sITNM1 = dt.Rows[0]["ITNM"]?.ToString();
            string sSPEC1 = dt.Rows[0]["SPEC"]?.ToString();
            double dTWGT1 = double.TryParse(dt.Rows[0]["TWGT"]?.ToString(), out double dRes1) ? double.Parse(dt.Rows[0]["TWGT"]?.ToString()) : 0 ;
            double dCOST1 = double.TryParse(dt.Rows[0]["COST"]?.ToString(), out double dRes2) ? double.Parse(dt.Rows[0]["COST"]?.ToString()) : 0 ;
            double dTAMT1 = double.TryParse(dt.Rows[0]["TAMT"]?.ToString(), out double dRes3) ? double.Parse(dt.Rows[0]["TAMT"]?.ToString()) : 0 ;
            double dTVAT1 = double.TryParse(dt.Rows[0]["TVAT"]?.ToString(), out double dRes4) ? double.Parse(dt.Rows[0]["TVAT"]?.ToString()) : 0 ;
            string sTRMK1 = dt.Rows[0]["TRMK"]?.ToString().Trim();

            string sMMDD2 = dt.Rows[1]["MMDD"]?.ToString();
            sMMDD2 = string.IsNullOrEmpty(sMMDD2) ? null : sMMDD2.Replace("-", "");
            string sITNM2 = dt.Rows[1]["ITNM"]?.ToString();
            string sSPEC2 = dt.Rows[1]["SPEC"]?.ToString();
            double dTWGT2 = double.TryParse(dt.Rows[1]["TWGT"]?.ToString(), out double dRes5) ? double.Parse(dt.Rows[1]["TWGT"]?.ToString()) : 0 ;
            double dCOST2 = double.TryParse(dt.Rows[1]["COST"]?.ToString(), out double dRes6) ? double.Parse(dt.Rows[1]["COST"]?.ToString()) : 0 ;
            double dTAMT2 = double.TryParse(dt.Rows[1]["TAMT"]?.ToString(), out double dRes7) ? double.Parse(dt.Rows[1]["TAMT"]?.ToString()) : 0 ;
            double dTVAT2 = double.TryParse(dt.Rows[1]["TVAT"]?.ToString(), out double dRes8) ? double.Parse(dt.Rows[1]["TVAT"]?.ToString()) : 0 ;
            string sTRMK2 = dt.Rows[1]["TRMK"]?.ToString().Trim();

            string sMMDD3 = dt.Rows[2]["MMDD"]?.ToString();
            sMMDD3 = string.IsNullOrEmpty(sMMDD3) ? null : sMMDD3.Replace("-", "");
            string sITNM3 = dt.Rows[2]["ITNM"]?.ToString();
            string sSPEC3 = dt.Rows[2]["SPEC"]?.ToString();
            double dTWGT3 = double.TryParse(dt.Rows[2]["TWGT"]?.ToString(), out double dRes9) ? double.Parse(dt.Rows[2]["TWGT"]?.ToString()) : 0 ; 
            double dCOST3 = double.TryParse(dt.Rows[2]["COST"]?.ToString(), out double dRes10) ? double.Parse(dt.Rows[2]["COST"]?.ToString()) : 0 ;
            double dTAMT3 = double.TryParse(dt.Rows[2]["TAMT"]?.ToString(), out double dRes11) ? double.Parse(dt.Rows[2]["TAMT"]?.ToString()) : 0 ;
            double dTVAT3 = double.TryParse(dt.Rows[2]["TVAT"]?.ToString(), out double dRes12) ? double.Parse(dt.Rows[2]["TVAT"]?.ToString()) : 0 ;
            string sTRMK3 = dt.Rows[2]["TRMK"]?.ToString().Trim();

            string sMMDD4 = dt.Rows[3]["MMDD"]?.ToString();
            sMMDD4 = string.IsNullOrEmpty(sMMDD4) ? null : sMMDD4.Replace("-", "");
            string sITNM4 = dt.Rows[3]["ITNM"]?.ToString();
            string sSPEC4 = dt.Rows[3]["SPEC"]?.ToString();
            double dTWGT4 = double.TryParse(dt.Rows[3]["TWGT"]?.ToString(), out double dRes13) ? double.Parse(dt.Rows[3]["TWGT"]?.ToString()) : 0 ;
            double dCOST4 = double.TryParse(dt.Rows[3]["COST"]?.ToString(), out double dRes14) ? double.Parse(dt.Rows[3]["COST"]?.ToString()) : 0 ;
            double dTAMT4 = double.TryParse(dt.Rows[3]["TAMT"]?.ToString(), out double dRes15) ? double.Parse(dt.Rows[3]["TAMT"]?.ToString()) : 0 ;
            double dTVAT4 = double.TryParse(dt.Rows[3]["TVAT"]?.ToString(), out double dRes16) ? double.Parse(dt.Rows[3]["TVAT"]?.ToString()) : 0 ;
            string sTRMK4 = dt.Rows[3]["TRMK"]?.ToString().Trim();

            //string sMMDD1 = GridViewRetr.GetRowCellValue(0, GridColMmDd)?.ToString();
            //sMMDD1 = string.IsNullOrEmpty(sMMDD1) ? null : sMMDD1.Replace("-", "").Substring(4, 4);
            //string sITNM1 = GridViewRetr.GetRowCellValue(0, GridColitNm)?.ToString();
            //string sSPEC1 = GridViewRetr.GetRowCellValue(0, GridColSpec)?.ToString();
            //string sTWGT1 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(0, GridColTQty)?.ToString()) ? null : GridViewRetr.GetRowCellValue(0, GridColTQty)?.ToString();
            //string sCOST1 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(0, GridColCost)?.ToString()) ? null : GridViewRetr.GetRowCellValue(0, GridColCost)?.ToString();
            //string sTAMT1 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(0, GridColTAmt)?.ToString()) ? null : GridViewRetr.GetRowCellValue(0, GridColTAmt)?.ToString();
            //string sTVAT1 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(0, GridColTVat)?.ToString()) ? null : GridViewRetr.GetRowCellValue(0, GridColTVat)?.ToString();
            //string sTRMK1 = GridViewRetr.GetRowCellValue(0, GridColTRmk)?.ToString();

            //string sMMDD2 = GridViewRetr.GetRowCellValue(1, GridColMmDd)?.ToString();
            //sMMDD2 = string.IsNullOrEmpty(sMMDD2) ? null : sMMDD2.Replace("-", "").Substring(4, 4);
            //string sITNM2 = GridViewRetr.GetRowCellValue(1, GridColitNm)?.ToString();
            //string sSPEC2 = GridViewRetr.GetRowCellValue(1, GridColSpec)?.ToString();
            //string sTWGT2 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(1, GridColTQty)?.ToString()) ? null : GridViewRetr.GetRowCellValue(1, GridColTQty)?.ToString();
            //string sCOST2 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(1, GridColCost)?.ToString()) ? null : GridViewRetr.GetRowCellValue(1, GridColCost)?.ToString();
            //string sTAMT2 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(1, GridColTAmt)?.ToString()) ? null : GridViewRetr.GetRowCellValue(1, GridColTAmt)?.ToString();
            //string sTVAT2 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(1, GridColTVat)?.ToString()) ? null : GridViewRetr.GetRowCellValue(1, GridColTVat)?.ToString();
            //string sTRMK2 = GridViewRetr.GetRowCellValue(1, GridColTRmk)?.ToString();

            //string sMMDD3 = GridViewRetr.GetRowCellValue(2, GridColMmDd)?.ToString();
            //sMMDD3 = string.IsNullOrEmpty(sMMDD3) ? null : sMMDD3.Replace("-", "").Substring(4, 4);
            //string sITNM3 = GridViewRetr.GetRowCellValue(2, GridColitNm)?.ToString();
            //string sSPEC3 = GridViewRetr.GetRowCellValue(2, GridColSpec)?.ToString();
            //string sTWGT3 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(2, GridColTQty)?.ToString()) ? null : GridViewRetr.GetRowCellValue(2, GridColTQty)?.ToString();
            //string sCOST3 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(2, GridColCost)?.ToString()) ? null : GridViewRetr.GetRowCellValue(2, GridColCost)?.ToString();
            //string sTAMT3 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(2, GridColTAmt)?.ToString()) ? null : GridViewRetr.GetRowCellValue(2, GridColTAmt)?.ToString();
            //string sTVAT3 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(2, GridColTVat)?.ToString()) ? null : GridViewRetr.GetRowCellValue(2, GridColTVat)?.ToString();
            //string sTRMK3 = GridViewRetr.GetRowCellValue(2, GridColTRmk)?.ToString();

            //string sMMDD4 = GridViewRetr.GetRowCellValue(3, GridColMmDd)?.ToString();
            //sMMDD4 = string.IsNullOrEmpty(sMMDD4) ? null : sMMDD4.Replace("-", "").Substring(4, 4);
            //string sITNM4 = GridViewRetr.GetRowCellValue(3, GridColitNm)?.ToString();
            //string sSPEC4 = GridViewRetr.GetRowCellValue(3, GridColSpec)?.ToString();
            //string sTWGT4 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(3, GridColTQty)?.ToString()) ? null : GridViewRetr.GetRowCellValue(3, GridColTQty)?.ToString();
            //string sCOST4 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(3, GridColCost)?.ToString()) ? null : GridViewRetr.GetRowCellValue(3, GridColCost)?.ToString();
            //string sTAMT4 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(3, GridColTAmt)?.ToString()) ? null : GridViewRetr.GetRowCellValue(3, GridColTAmt)?.ToString();
            //string sTVAT4 = string.IsNullOrEmpty(GridViewRetr.GetRowCellValue(3, GridColTVat)?.ToString()) ? null : GridViewRetr.GetRowCellValue(3, GridColTVat)?.ToString();
            //string sTRMK4 = GridViewRetr.GetRowCellValue(3, GridColTRmk)?.ToString();

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                if (string.IsNullOrEmpty(sTaxNo))
                {
                    //TAXNO 채번
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" SELECT CONCAT('TX', DATE_FORMAT(NOW(), '%Y%m%d') , LPAD(IFNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1, 5, '0')) AS MAX_VAL ");
                    //strSql.AppendLine("   FROM TAXF A  ");
                    //strSql.AppendLine("  WHERE A.TDATE = @TDATE ");
                    #endregion

                    strSql.AppendLine("SELECT CONCAT('TX', CONVERT(VARCHAR(8), CONVERT(DATE, @TDATE), 112)                    ");
                    strSql.AppendLine("               , CAST(REPLICATE(0, 5 - LEN(ISNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1)) AS VARCHAR) ");
                    strSql.AppendLine("               + CAST(ISNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1 AS VARCHAR)) AS MAX_VAL            ");
                    strSql.AppendLine("  FROM TAXF A                                                                                ");
                    strSql.AppendLine(" WHERE A.TDATE = @TDATE                                                                ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@TDATE", DateTime.Today.ToString("yyyy-MM-dd"));
                    sTaxNo = cmd.ExecuteScalar()?.ToString();
                }

                string[] sArrDate = GetDealTerm(GridViewRetr);

                //Dictionary<string, string> dicParams = new Dictionary<string, string>();

                //dicParams.Add("TAXNO", sTaxNo);
                //dicParams.Add("PDAT1", sArrDate[0]);
                //dicParams.Add("PDAT2", sArrDate[1]);
                //dicParams.Add("TDATE", sTDate);
                //dicParams.Add("CVCOD", sCvCOd);
                //dicParams.Add("TGUBN", sTGubn);
                //dicParams.Add("PAYGB", sPayGb);
                //dicParams.Add("TAXGU", sTaxGu);
                //dicParams.Add("AUTGB", "2");
                //dicParams.Add("RK", sRK);

                //dicParams.Add("MMDD1", sMMDD1);
                //dicParams.Add("ITNM1", sITNM1);
                //dicParams.Add("SPEC1", sSPEC1);
                //dicParams.Add("TWGT1", sTWGT1);
                //dicParams.Add("COST1", sCOST1);
                //dicParams.Add("TAMT1", sTAMT1);
                //dicParams.Add("TVAT1", sTVAT1);
                //dicParams.Add("TRMK1", sTRMK1);

                //dicParams.Add("MMDD2", sMMDD2);
                //dicParams.Add("ITNM2", sITNM2);
                //dicParams.Add("SPEC2", sSPEC2);
                //dicParams.Add("TWGT2", sTWGT2);
                //dicParams.Add("COST2", sCOST2);
                //dicParams.Add("TAMT2", sTAMT2);
                //dicParams.Add("TVAT2", sTVAT2);
                //dicParams.Add("TRMK2", sTRMK2);

                //dicParams.Add("MMDD3", sMMDD3);
                //dicParams.Add("ITNM3", sITNM3);
                //dicParams.Add("SPEC3", sSPEC3);
                //dicParams.Add("TWGT3", sTWGT3);
                //dicParams.Add("COST3", sCOST3);
                //dicParams.Add("TAMT3", sTAMT3);
                //dicParams.Add("TVAT3", sTVAT3);
                //dicParams.Add("TRMK3", sTRMK3);

                //dicParams.Add("MMDD4", sMMDD4);
                //dicParams.Add("ITNM4", sITNM4);
                //dicParams.Add("SPEC4", sSPEC4);
                //dicParams.Add("TWGT4", sTWGT4);
                //dicParams.Add("COST4", sCOST4);
                //dicParams.Add("TAMT4", sTAMT4);
                //dicParams.Add("TVAT4", sTVAT4);
                //dicParams.Add("TRMK4", sTRMK4);

                //dicParams.Add("CUSER", FmMainToolBar2.UserID);
                //dicParams.Add("MUSER", FmMainToolBar2.UserID);

                strSql.Clear();
                #region mariaDB
                //strSql.AppendLine(" INSERT INTO TAXF ");
                //strSql.AppendLine("           ( TAXNO, TDATE, PDAT1, PDAT2, CVCOD ");
                //strSql.AppendLine("           , TGUBN, PAYGB, TAXGU, AUTGB, RK ");
                //strSql.AppendLine("           , MMDD1, ITNM1, SPEC1, TWGT1, COST1, TAMT1, TVAT1, TRMK1 ");
                //strSql.AppendLine("           , MMDD2, ITNM2, SPEC2, TWGT2, COST2, TAMT2, TVAT2, TRMK2 ");
                //strSql.AppendLine("           , MMDD3, ITNM3, SPEC3, TWGT3, COST3, TAMT3, TVAT3, TRMK3 ");
                //strSql.AppendLine("           , MMDD4, ITNM4, SPEC4, TWGT4, COST4, TAMT4, TVAT4, TRMK4 ");
                //strSql.AppendLine("           , CDATE, CUSER ) ");
                //strSql.AppendLine("     VALUES( @TAXNO, @TDATE, @PDAT1, @PDAT2, @CVCOD ");
                //strSql.AppendLine("           , @TGUBN, @PAYGB, @TAXGU, @AUTGB, @RK ");
                //strSql.AppendLine("           , @MMDD1, @ITNM1, @SPEC1, @TWGT1, @COST1, @TAMT1, @TVAT1, @TRMK1 ");
                //strSql.AppendLine("           , @MMDD2, @ITNM2, @SPEC2, @TWGT2, @COST2, @TAMT2, @TVAT2, @TRMK2 ");
                //strSql.AppendLine("           , @MMDD3, @ITNM3, @SPEC3, @TWGT3, @COST3, @TAMT3, @TVAT3, @TRMK3 ");
                //strSql.AppendLine("           , @MMDD4, @ITNM4, @SPEC4, @TWGT4, @COST4, @TAMT4, @TVAT4, @TRMK4 ");
                //strSql.AppendLine("           , NOW(), @CUSER ) ");
                //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                //strSql.AppendLine("             TAXNO = @TAXNO, TDATE = @TDATE, PDAT1 = @PDAT1, PDAT2 = @PDAT2, CVCOD = @CVCOD ");
                //strSql.AppendLine("           , TGUBN = @TGUBN, PAYGB = @PAYGB, TAXGU = @TAXGU, AUTGB = @AUTGB, RK = @RK ");
                //strSql.AppendLine("           , MMDD1 = @MMDD1, ITNM1 = @ITNM1, SPEC1 = @SPEC1, TWGT1 = @TWGT1, COST1 = @COST1, TAMT1 = @TAMT1, TVAT1 = @TVAT1, TRMK1 = @TRMK1 ");
                //strSql.AppendLine("           , MMDD2 = @MMDD2, ITNM2 = @ITNM2, SPEC2 = @SPEC2, TWGT2 = @TWGT2, COST2 = @COST2, TAMT2 = @TAMT2, TVAT2 = @TVAT2, TRMK2 = @TRMK2 ");
                //strSql.AppendLine("           , MMDD3 = @MMDD3, ITNM3 = @ITNM3, SPEC3 = @SPEC3, TWGT3 = @TWGT3, COST3 = @COST3, TAMT3 = @TAMT3, TVAT3 = @TVAT3, TRMK3 = @TRMK3 ");
                //strSql.AppendLine("           , MMDD4 = @MMDD4, ITNM4 = @ITNM4, SPEC4 = @SPEC4, TWGT4 = @TWGT4, COST4 = @COST4, TAMT4 = @TAMT4, TVAT4 = @TVAT4, TRMK4 = @TRMK4 ");
                //strSql.AppendLine("           , CDATE = NOW(), CUSER = @CUSER ");
                #endregion

                strSql.AppendLine("IF EXISTS(SELECT* FROM TAXF WHERE TAXNO = '" + sTaxNo + "')                                                                                            ");
                strSql.AppendLine("   BEGIN                                                                                                                                     ");
                strSql.AppendLine("         UPDATE TAXF                                                                                                                         ");
                strSql.AppendLine("            SET TDATE = '" + sTDate + "', PDAT1 = '"+ sArrDate[0] + "', PDAT2 = '" + sArrDate[1] + "', CVCOD = "+ sCvCOd + "                                                               ");
                strSql.AppendLine("	          , TGUBN = '"+ sTGubn + "', PAYGB = '"+ sPayGb + "', TAXGU = '"+ sTaxGu + "', AUTGB = '2', RK = '"+ sRK + "'                                                        ");
                strSql.AppendLine("	          , MMDD1 = '"+ sMMDD1 + "', ITNM1 = '"+sITNM1+"', SPEC1 = '"+sSPEC1+"', TWGT1 = "+dTWGT1+", COST1 = "+dCOST1+", TAMT1 = "+dTAMT1+", TVAT1 ="+dTVAT1+", TRMK1 = '"+sTRMK1+"'  ");
                strSql.AppendLine("	          , MMDD2 = '"+ sMMDD2 + "', ITNM2 = '"+sITNM2+"', SPEC2 = '"+sSPEC2+"', TWGT2 = "+dTWGT2+", COST2 = "+dCOST2+", TAMT2 = "+dTAMT2+", TVAT2 ="+dTVAT2+", TRMK2 = '"+sTRMK2+"'  ");
                strSql.AppendLine("	          , MMDD3 = '"+ sMMDD3 + "', ITNM3 = '"+sITNM3+"', SPEC3 = '"+sSPEC3+"', TWGT3 = "+dTWGT3+", COST3 = "+dCOST3+", TAMT3 = "+dTAMT3+", TVAT3 ="+dTVAT3+", TRMK3 = '"+sTRMK3+"'  ");
                strSql.AppendLine("	          , MMDD4 = '"+ sMMDD4 + "', ITNM4 = '"+sITNM4+"', SPEC4 = '"+sSPEC4+"', TWGT4 = "+dTWGT4+", COST4 = "+dCOST4+", TAMT4 = "+dTAMT4+", TVAT4 ="+dTVAT4+", TRMK4 = '"+sTRMK4+"'  ");
                strSql.AppendLine("	          , MDATE = CONVERT(VARCHAR(19), GETDATE(), 20), MUSER = "+ FmMainToolBar2.UserID + "                                                                     ");
                strSql.AppendLine("          WHERE TAXNO = '" + sTaxNo + "'                                                                                                               ");
                strSql.AppendLine("     END                                                                                                                                     ");
                strSql.AppendLine("ELSE                                                                                                                                         ");
                strSql.AppendLine("   BEGIN                                                                                                                                     ");
                strSql.AppendLine("         INSERT INTO TAXF                                                                                                                    ");
                strSql.AppendLine("           (TAXNO, TDATE, PDAT1, PDAT2, CVCOD                                                                                                ");
                strSql.AppendLine("           , TGUBN, PAYGB, TAXGU, AUTGB, RK                                                                                                  ");
                strSql.AppendLine("           , MMDD1, ITNM1, SPEC1, TWGT1, COST1, TAMT1, TVAT1, TRMK1                                                                          ");
                strSql.AppendLine("           , MMDD2, ITNM2, SPEC2, TWGT2, COST2, TAMT2, TVAT2, TRMK2                                                                          ");
                strSql.AppendLine("           , MMDD3, ITNM3, SPEC3, TWGT3, COST3, TAMT3, TVAT3, TRMK3                                                                          ");
                strSql.AppendLine("           , MMDD4, ITNM4, SPEC4, TWGT4, COST4, TAMT4, TVAT4, TRMK4                                                                          ");
                strSql.AppendLine("           , CDATE, CUSER )                                                                                                                  ");
                strSql.AppendLine("     VALUES('" + sTaxNo + "', '" + sTDate + "', '" + sArrDate[0] + "', '" + sArrDate[1] + "', "+ sCvCOd + "                                                                                          ");
                strSql.AppendLine("           , '" + sTGubn + "', '" + sPayGb + "', '"+ sTaxGu + "', '2', '"+ sRK + "'                                                                                             ");
                strSql.AppendLine("           , '"+ sMMDD1 + "', '"+sITNM1+"', '"+sSPEC1+"', "+dTWGT1+", "+dCOST1+", "+dTAMT1+", "+dTVAT1+", '"+sTRMK1+"'                                                                  ");
                strSql.AppendLine("           , '"+ sMMDD2 + "', '"+sITNM2+"', '"+sSPEC2+"', "+dTWGT2+", "+dCOST2+", "+dTAMT2+", "+dTVAT2+", '"+sTRMK2+"'                                                                  ");
                strSql.AppendLine("           , '"+ sMMDD3 + "', '"+sITNM3+"', '"+sSPEC3+"', "+dTWGT3+", "+dCOST3+", "+dTAMT3+", "+dTVAT3+", '"+sTRMK3+"'                                                                  ");
                strSql.AppendLine("           , '"+ sMMDD4 + "', '"+sITNM4+ "','"+sSPEC4+ "',"+dTWGT4+", "+dCOST4+", "+dTAMT4+", "+dTVAT4+", '"+sTRMK4+"'                                                                  ");
                strSql.AppendLine("           , CONVERT(VARCHAR(19), GETDATE(), 20), "+ FmMainToolBar2.UserID + ")                                                                                    ");
                strSql.AppendLine("     END                                                                                                                                     ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                //cmd.Parameters.Clear();
                //foreach (KeyValuePair<string, string> param in dicParams)
                //{
                //    cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                //}
                cmd.ExecuteNonQuery();
                //cmd.Parameters.Clear();

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                DialogResult = DialogResult.OK;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private bool ValidateGridViewEssentialData(GridView view)
        {
            int iNotInputCnt = 0;
            bool bYn = true;

            for (int i = 0; i < view.RowCount; i++)
            {
                DataRow row = view.GetDataRow(i);

                string sMMDD = row["MMDD"]?.ToString();
                string sITNM = row["ITNM"]?.ToString();
                string sSPEC = row["SPEC"]?.ToString();
                double dTWGT = string.IsNullOrEmpty(row["TWGT"]?.ToString()) ? 0 : Convert.ToDouble(row["TWGT"]);
                double dCOST = string.IsNullOrEmpty(row["COST"]?.ToString()) ? 0 : Convert.ToDouble(row["COST"]);
                string sRMK = row["TRMK"]?.ToString();

                //전체가 입력되지 않을 경우 카운트 증가;
                if (string.IsNullOrEmpty(sMMDD) && string.IsNullOrEmpty(sITNM) && string.IsNullOrEmpty(sSPEC)
                    && dTWGT == 0 && dCOST == 0 && string.IsNullOrEmpty(sRMK))
                {
                    iNotInputCnt++;
                    continue;
                }
                //필수입력이 아닌 컬럼에 값이 존재하지만 필수입력인 컬럼의 전체가 입력되지 않을 경우 노란색
                else if ((string.IsNullOrEmpty(sMMDD) && string.IsNullOrEmpty(sITNM) && dTWGT == 0 && dCOST == 0) && (!string.IsNullOrEmpty(sSPEC) || !string.IsNullOrEmpty(sRMK)))
                {
                    if (string.IsNullOrEmpty(sMMDD))
                    {
                        XtraMessageBox.Show("월일을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColMmDd;
                        //bYn = false;
                        return false;
                    }
                    else if (string.IsNullOrEmpty(sITNM))
                    {
                        XtraMessageBox.Show("품명을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColitNm;
                        //bYn = false;
                        return false;
                    }
                    else if(dTWGT <= 0)
                    {
                        XtraMessageBox.Show("수량을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTQty;
                        //bYn = false;
                        return false;
                    }
                    else if(dCOST <= 0)
                    {
                        XtraMessageBox.Show("단가를 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColCost;
                        //bYn = false;
                        return false;
                    }
                }
                else if (string.IsNullOrEmpty(sMMDD) || string.IsNullOrEmpty(sITNM) || dTWGT == 0 || dCOST == 0)
                {
                    if (string.IsNullOrEmpty(sMMDD))
                    {
                        XtraMessageBox.Show("월일을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColMmDd;
                        //bYn = false;
                        return false;
                    }
                    else if (string.IsNullOrEmpty(sITNM))
                    {
                        XtraMessageBox.Show("품명을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColitNm;
                        //bYn = false;
                        return false;
                    }
                    else if (dTWGT <= 0)
                    {
                        XtraMessageBox.Show("수량을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTQty;
                        //bYn = false;
                        return false;
                    }
                    else if (dCOST <= 0)
                    {
                        XtraMessageBox.Show("단가를 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColCost;
                        //bYn = false;
                        return false;
                    }
                }
            }

            if(view.RowCount == iNotInputCnt)
            {
                XtraMessageBox.Show("상세내역을 한건 이상 입력하여야 합니다.");
                bYn = false;
            }

            return bYn;
        }

        private DataTable ReArrangeRowSequence(DataTable dt)
        {
            DataTable dtCopy = new DataTable();
            dtCopy = dt.Clone();
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];

                string sMMDD = row["MMDD"]?.ToString();
                string sITNM = row["ITNM"]?.ToString();
                string sSPEC = row["SPEC"]?.ToString();
                double dTWGT = string.IsNullOrEmpty(row["TWGT"]?.ToString()) ? 0 : Convert.ToDouble(row["TWGT"]);
                double dCOST = string.IsNullOrEmpty(row["COST"]?.ToString()) ? 0 : Convert.ToDouble(row["COST"]);
                string sRMK = row["TRMK"]?.ToString();

                //전체가 입력되지 않을 경우 카운트 증가;
                if (string.IsNullOrEmpty(sMMDD) && string.IsNullOrEmpty(sITNM) && string.IsNullOrEmpty(sSPEC)
                    && dTWGT == 0 && dCOST == 0 && string.IsNullOrEmpty(sRMK))
                {
                    continue;
                }
                //필수입력이 아닌 컬럼에 값이 존재하지만 필수입력인 컬럼의 전체가 입력되지 않을 경우 노란색
                else if ((string.IsNullOrEmpty(sMMDD) && string.IsNullOrEmpty(sITNM) && dTWGT == 0 && dCOST == 0) && (!string.IsNullOrEmpty(sSPEC) || !string.IsNullOrEmpty(sRMK)))
                {
                    continue;
                }
                else if (string.IsNullOrEmpty(sMMDD) || string.IsNullOrEmpty(sITNM) || dTWGT == 0 || dCOST == 0)
                {
                    continue;
                }
                else
                {
                    dtCopy.ImportRow(row);
                    //dtCopy.Rows.Add(row);
                }
            }

            for (int i = dtCopy.Rows.Count; i <= 4; i++)
            {
                DataRow row = dtCopy.NewRow();
                row["TWGT"] = DBNull.Value;
                row["COST"] = DBNull.Value;
                row["TAMT"] = DBNull.Value;
                row["TVAT"] = DBNull.Value;
                dtCopy.Rows.Add(row);
            }

            return dtCopy;
        }

        private string[] GetDealTerm(GridView view)
        {
            string[] sArr = new string[2];
            
            Dictionary<int, DateTime> dicTemp = new Dictionary<int, DateTime>();
            for(int i = 0; i < view.RowCount; i++)
            {
                string sMMDD = view.GetRowCellValue(i, GridColMmDd)?.ToString();
                if (string.IsNullOrEmpty(sMMDD))
                    continue;

                dicTemp.Add(i, Convert.ToDateTime(sMMDD));
            }
            
            sArr[0] = dicTemp.Values.Min().ToString("yyyy-MM-dd");
            sArr[1] = dicTemp.Values.Max().ToString("yyyy-MM-dd");

            return sArr;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void AC18001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }

        //초기 LookupEdit 세팅
        private void SetLookup()
        {
            DataTable dtTGubn = GetLookUpData("1", "Y", ""); //청구영수구분
            LkupTGubn.Properties.ValueMember = "CD";
            LkupTGubn.Properties.DisplayMember = "NM";
            LkupTGubn.Properties.DataSource = dtTGubn;

            DataTable dtPayGb = GetLookUpData("2", "Y", ""); //지급구분
            LkupPayGb.Properties.ValueMember = "CD";
            LkupPayGb.Properties.DisplayMember = "NM";
            LkupPayGb.Properties.DataSource = dtPayGb;

            DataTable dtTaxGu = GetLookUpData("3", "Y", ""); //발행구분
            LkupTaxGu.Properties.ValueMember = "CD";
            LkupTaxGu.Properties.DisplayMember = "NM";
            LkupTaxGu.Properties.DataSource = dtTaxGu;
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
                strSql.AppendLine("  WHERE CD_GB = 'AC18001_01' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine("  WHERE CD_GB = 'AC18001_02' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine("  WHERE CD_GB = 'AC18001_03' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY CASE WHEN CD = '' THEN '100' ELSE CD END");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        public DataRow DrDealerInfo;
        private void BtnCvNam_KeyDown(object sender, KeyEventArgs e)
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
                        btnEdit.EditValue = dt.Rows[0]["DEALER_NM"];
                        TxtCvCod.EditValue = dt.Rows[0]["DEALER_CD"];
                        TxtIdtNo.EditValue = dt.Rows[0]["IDT_NO"];
                        TxtAddr.EditValue = dt.Rows[0]["ADDR"];
                        TxtRepNm.EditValue = dt.Rows[0]["REP_NM"];
                        TxtUpTae.EditValue = dt.Rows[0]["BIZ_NM"];
                        TxtJongK.EditValue = dt.Rows[0]["TYPE_NM"];
                    }
                    else
                    {
                        AC02001F03 frm = new AC02001F03();
                        frm.P_AC18001F02 = this;
                        frm.DealerCd = sDealerCd;
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            btnEdit.EditValue = DrDealerInfo["DEALER_NM"];
                            TxtCvCod.EditValue = DrDealerInfo["DEALER_CD"];
                            TxtIdtNo.EditValue = DrDealerInfo["IDT_NO"];
                            TxtAddr.EditValue = DrDealerInfo["ADDR"];
                            TxtRepNm.EditValue = DrDealerInfo["REP_NM"];
                            TxtUpTae.EditValue = DrDealerInfo["BIZ_NM"];
                            TxtJongK.EditValue = DrDealerInfo["TYPE_NM"];
                        }
                    }
                }
                else
                {
                    btnEdit.EditValue = null;
                    TxtCvCod.EditValue = null;
                }

                DrDealerInfo = null;
            }
        }

        private void BtnCvNam_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            AC02001F03 frm = new AC02001F03();
            frm.P_AC18001F02 = this;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnCvNam.EditValue = DrDealerInfo["DEALER_NM"];
                TxtCvCod.EditValue = DrDealerInfo["DEALER_CD"];
                TxtIdtNo.EditValue = DrDealerInfo["IDT_NO"];
                TxtAddr.EditValue = DrDealerInfo["ADDR"];
                TxtRepNm.EditValue = DrDealerInfo["REP_NM"];
                TxtUpTae.EditValue = DrDealerInfo["BIZ_NM"];
                TxtJongK.EditValue = DrDealerInfo["TYPE_NM"];
                DrDealerInfo = null;
            }
        }

        private DataTable GetDealerInfo(string sDealerCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.DEALER_CD  ");
            strSql.AppendLine(" 	 , CAST(A.DEALER_NM AS CHAR) AS DEALER_NM ");
            strSql.AppendLine(" 	 , A.IDT_NO  ");
            strSql.AppendLine(" 	 , A.DEALER_GB  ");
            strSql.AppendLine("      , CASE WHEN A.ADDR IS NULL OR REPLACE(A.ADDR, ' ', '') = '' THEN DTL_ADDR ELSE ADDR END AS ADDR ");
            strSql.AppendLine("      , A.REP_NM ");
            strSql.AppendLine("      , A.BIZ_NM ");
            strSql.AppendLine("      , A.TYPE_NM ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A  ");
            strSql.AppendLine("  WHERE (A.DEALER_CD LIKE '" + sDealerCd + "' ");
            strSql.AppendLine("     OR A.DEALER_NM LIKE '%" + sDealerCd + "%' ");
            strSql.AppendLine("     OR A.INITIAL_NM LIKE '%" + sDealerCd + "%') ");
            strSql.AppendLine("    AND A.EOB_YN = 'N' ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            if (e.RowHandle < 0)
                return;

            DataRow row = GridViewRetr.GetDataRow(e.RowHandle);
            string sMMDD = row["MMDD"]?.ToString();
            string sITNM = row["ITNM"]?.ToString();
            string sSPEC = row["SPEC"]?.ToString();
            double dTWGT = string.IsNullOrEmpty(row["TWGT"]?.ToString()) ? 0 : Convert.ToDouble(row["TWGT"]);
            double dCOST = string.IsNullOrEmpty(row["COST"]?.ToString()) ? 0 : Convert.ToDouble(row["COST"]);
            string sRMK = row["TRMK"]?.ToString();

            //전체가 입력되지 않을 경우 return;
            if(string.IsNullOrEmpty(sMMDD) && string.IsNullOrEmpty(sITNM) && string.IsNullOrEmpty(sSPEC) 
                && dTWGT == 0 && dCOST == 0 && string.IsNullOrEmpty(sRMK))
            {
                return;
            }
            //필수입력이 아닌 컬럼에 값이 존재하지만 필수입력인 컬럼의 전체가 입력되지 않을 경우 노란색
            else if((string.IsNullOrEmpty(sMMDD) && string.IsNullOrEmpty(sITNM) && dTWGT == 0 && dCOST == 0) && (!string.IsNullOrEmpty(sSPEC) || !string.IsNullOrEmpty(sRMK)))
            {
                e.Appearance.BackColor = Color.Yellow;
            }
            else if (string.IsNullOrEmpty(sMMDD) || string.IsNullOrEmpty(sITNM) || dTWGT == 0 || dCOST == 0)
            {
                e.Appearance.BackColor = Color.Yellow;
            }
            //필수입력 컬럼이 전부입력되었을 경우 초록색
            else if(!string.IsNullOrEmpty(sMMDD) && !string.IsNullOrEmpty(sITNM) && dTWGT > 0 && dCOST > 0)
            {
                e.Appearance.BackColor = Color.PaleGreen;
            }

        }

        private void RepoTxtTQty_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridColTQty, txt.EditValue);
        }

        private void RepoTxtTQty_Leave(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            double dTWGT = string.IsNullOrEmpty(txt.EditValue?.ToString()) ? 0 : Convert.ToDouble(txt.EditValue);
            double dCOST = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColCost)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColCost));

            GridViewRetr.SetFocusedRowCellValue(GridColTAmt, dTWGT * dCOST);
            GridViewRetr.SetFocusedRowCellValue(GridColTVat, (dTWGT * dCOST) * 0.1);
            GridViewRetr.SetFocusedRowCellValue(GridColTotAmt, (dTWGT * dCOST) + ((dTWGT * dCOST) * 0.1));
        }

        private void RepoTxtCost_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridColCost, txt.EditValue);
        }
        
        private void RepoTxtCost_Leave(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            double dCOST = string.IsNullOrEmpty(txt.EditValue?.ToString()) ? 0 : Convert.ToDouble(txt.EditValue);
            double dTWGT = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColTQty)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColTQty));

            GridViewRetr.SetFocusedRowCellValue(GridColTAmt, dTWGT * dCOST);
            GridViewRetr.SetFocusedRowCellValue(GridColTVat, (dTWGT * dCOST) * 0.1);
            GridViewRetr.SetFocusedRowCellValue(GridColTotAmt, (dTWGT * dCOST) + ((dTWGT * dCOST) * 0.1));
        }

        private void GridViewRetr_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            
        }

        private void LkupTaxGu_EditValueChanged(object sender, EventArgs e)
        {
            string sTaxGu = LkupTaxGu.EditValue?.ToString();
            string sTGubn = LkupTGubn.EditValue?.ToString();
            string sPayGb = LkupPayGb.EditValue?.ToString();

            if (string.IsNullOrEmpty(sTaxGu))
                return;

            if (sTaxGu.Equals("P")) //매입
            {
                LkupTGubn.EditValue = "2";
                LkupPayGb.EditValue = "4";
            }
            else if(sTaxGu.Equals("S")) //매출
            {
                LkupTGubn.EditValue = "1";
                LkupPayGb.EditValue = "5";
            }
        }

        private void RepoDateEditMmDd_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = sender as DateEdit;

            if(DateTime.TryParse(dateEdit.EditValue?.ToString(), out DateTime result))
            {
                DateTime dt = DateTime.Parse(dateEdit.EditValue?.ToString());

                GridViewRetr.SetFocusedRowCellValue(GridColMmDd, dt.ToString("MM-dd"));
            }
        }
    }
}