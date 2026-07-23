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
 * 작성일자 : 2021-03-29
 * 작성자   : 고혜성
 * 작성내용 : 전표에서 세금계산서 발행할 수 있도록 세팅
 */
namespace AccAdm
{
    public partial class AC02001F08 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F08()
        {
            InitializeComponent();
        }
        public string _TDATE = string.Empty;
        public string _ATGUB = string.Empty;
        public string _SEQNO = string.Empty;

        private void AC02001F08_Load(object sender, EventArgs e)
        {
            Init();
        }

        private void AC02001F08_Shown(object sender, EventArgs e)
        {
            
        }
        
        private void Init()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("TDATE", _TDATE);
                dicParams.Add("ATGUB", _ATGUB);
                dicParams.Add("SEQNO", _SEQNO);

                DataTable dtSlip = GetSlipInfo(dicParams);
                GridSlip.DataSource = dtSlip;

                string sTaxNo = string.Empty;
                if(dtSlip.Rows.Count > 0)
                {
                    sTaxNo = dtSlip.Rows[0]["TAXNO"]?.ToString();
                }

                DataTable dtSlipInfo = GetTaxInfo(sTaxNo);
                if (dtSlipInfo.Rows.Count > 0)
                {
                    TxtTaxNo.EditValue = dtSlipInfo.Rows[0]["TAXNO"];
                    DateEditTDate.EditValue = dtSlipInfo.Rows[0]["TDATE"];
                    LkupTaxGu.EditValue = dtSlipInfo.Rows[0]["TAXGU"];
                    LkupTGubn.EditValue = dtSlipInfo.Rows[0]["TGUBN"];
                    LkupPayGb.EditValue = dtSlipInfo.Rows[0]["PAYGB"];
                    TxtCvCod.EditValue = dtSlipInfo.Rows[0]["CVCOD"];
                    BtnCvNam.EditValue = dtSlipInfo.Rows[0]["DEALER_NM"];
                    TxtIdtNo.EditValue = dtSlipInfo.Rows[0]["IDT_NO"];
                    TxtAddr.EditValue = dtSlipInfo.Rows[0]["ADDR"];
                    TxtRepNm.EditValue = dtSlipInfo.Rows[0]["REP_NM"];
                    TxtUpTae.EditValue = dtSlipInfo.Rows[0]["BIZ_NM"];
                    TxtJongK.EditValue = dtSlipInfo.Rows[0]["TYPE_NM"];
                    TxtTotAmt.EditValue = dtSlipInfo.Rows[0]["TAMT"];
                    TxtTotVat.EditValue = dtSlipInfo.Rows[0]["TVAT"];
                    MemoRmk.EditValue = dtSlipInfo.Rows[0]["RK"];
                }
                
                GridRetr.DataSource = GetTaxDetInfo(sTaxNo);

                if (string.IsNullOrEmpty(sTaxNo))
                {
                    LkupTaxGu.SelectAll();
                    LkupTaxGu.Focus();
                }
                else
                {
                    GridViewRetr.Focus();
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show("[Init Error] " + ex.Message);
            }
            
        }

        private DataTable GetSlipInfo(Dictionary<string, string> dicParams)
        {

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendFormat("\r\n ");
            strSql.AppendFormat("\r\n SELECT A1.TDATE ");
            strSql.AppendFormat("\r\n      , DATE_FORMAT(A1.TDATE, '%Y-%m-%d') AS TDATE2 ");
            strSql.AppendFormat("\r\n      , A1.ATGUB ");
            strSql.AppendFormat("\r\n      , B1.COM_NM AS ATGUB2 ");
            strSql.AppendFormat("\r\n      , A1.SEQNO ");
            strSql.AppendFormat("\r\n      , A1.LINNO ");
            strSql.AppendFormat("\r\n      , A1.ACCOD ");
            strSql.AppendFormat("\r\n      , A2.ACNAM ");
            strSql.AppendFormat("\r\n      , A1.CVCOD ");
            strSql.AppendFormat("\r\n      , A3.DEALER_NM ");
            strSql.AppendFormat("\r\n      , A1.ATEXT ");
            strSql.AppendFormat("\r\n      , A1.REF1 ");
            strSql.AppendFormat("\r\n      , A1.REF2 ");
            strSql.AppendFormat("\r\n      , A1.REF3 ");
            strSql.AppendFormat("\r\n      , A1.TAXNO ");
            strSql.AppendFormat("\r\n      , A1.APVYN ");
            strSql.AppendFormat("\r\n   FROM ACTRAN A1 ");
            strSql.AppendFormat("\r\n   LEFT JOIN ACMSTF A2  ");
            strSql.AppendFormat("\r\n     ON A1.ACCOD = A2.ACCOD ");
            strSql.AppendFormat("\r\n   LEFT JOIN ACC_DEALER_CD A3 ");
            strSql.AppendFormat("\r\n     ON A1.CVCOD = A3.DEALER_CD ");
            strSql.AppendFormat("\r\n   LEFT JOIN COM_BASE_CD B1 ");
            strSql.AppendFormat("\r\n     ON A1.ATGUB = B1.COM_CD ");
            strSql.AppendFormat("\r\n    AND B1.CD_GB = 'AC02001_01' ");
            strSql.AppendFormat("\r\n  WHERE A1.TDATE = '{0}' ", dicParams["TDATE"]);
            strSql.AppendFormat("\r\n    AND A1.ATGUB = '{0}' ", dicParams["ATGUB"]);
            strSql.AppendFormat("\r\n    AND A1.SEQNO = '{0}' ", dicParams["SEQNO"]);
            strSql.AppendFormat("\r\n  ORDER BY A1.TDATE, A1.ATGUB, A1.SEQNO, A1.LINNO ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        }

        /// <summary>
        ///    계산서 기본정보 가져오기
        /// </summary>
        /// <param name="TaxNo">계산서번호</param>
        private DataTable GetTaxInfo(string TaxNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendFormat("\r\n ");
            strSql.AppendFormat("\r\n SELECT A.TAXNO ");
            strSql.AppendFormat("\r\n      , A.TDATE ");
            strSql.AppendFormat("\r\n      , A.TAXGU ");
            strSql.AppendFormat("\r\n      , A.TGUBN ");
            strSql.AppendFormat("\r\n      , A.PAYGB ");
            strSql.AppendFormat("\r\n      , A.CVCOD ");
            strSql.AppendFormat("\r\n      , B.DEALER_NM ");
            strSql.AppendFormat("\r\n      , B.IDT_NO ");
            strSql.AppendFormat("\r\n      , B.ADDR ");
            strSql.AppendFormat("\r\n      , B.REP_NM ");
            strSql.AppendFormat("\r\n      , B.BIZ_NM ");
            strSql.AppendFormat("\r\n      , B.TYPE_NM ");
            strSql.AppendFormat("\r\n      , IFNULL(TAMT1, 0) + IFNULL(TAMT2, 0) + IFNULL(TAMT3, 0) + IFNULL(TAMT4, 0) AS TAMT ");
            strSql.AppendFormat("\r\n      , IFNULL(TVAT1, 0) + IFNULL(TVAT2, 0) + IFNULL(TVAT3, 0) + IFNULL(TVAT4, 0) AS TVAT ");
            strSql.AppendFormat("\r\n      , A.RK ");
            strSql.AppendFormat("\r\n   FROM TAXF A ");
            strSql.AppendFormat("\r\n   LEFT JOIN ACC_DEALER_CD B  ");
            strSql.AppendFormat("\r\n     ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendFormat("\r\n  WHERE A.TAXNO = '{0}' ", TaxNo);

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        /// <summary>
        ///    계산서 상세내역 가져오기
        /// </summary>
        /// <param name="TaxNo">계산서번호</param>
        private DataTable GetTaxDetInfo(string TaxNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendFormat("\r\n ");
            strSql.AppendFormat("\r\n ");
            strSql.AppendFormat("\r\n  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD1, 2), '-', RIGHT(MMDD1, 2)) AS MMDD, ITNM1 AS ITNM, SPEC1 AS SPEC, TWGT1 AS TWGT ");
            strSql.AppendFormat("\r\n       , COST1 AS COST, TAMT1 AS TAMT, TVAT1 AS TVAT, TRMK1 AS TRMK ");
            strSql.AppendFormat("\r\n       , TAMT1 + TVAT1 AS TOT_AMT ");
            strSql.AppendFormat("\r\n    FROM TAXF X1 ");
            strSql.AppendFormat("\r\n   WHERE TAXNO = '{0}' ", TaxNo);
            strSql.AppendFormat("\r\n   UNION ALL ");
            strSql.AppendFormat("\r\n  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD2, 2), '-', RIGHT(MMDD2, 2)), ITNM2, SPEC2, TWGT2 ");
            strSql.AppendFormat("\r\n       , COST2, TAMT2, TVAT2, TRMK2 ");
            strSql.AppendFormat("\r\n       , TAMT2 + TVAT2 AS TOT_AMT ");
            strSql.AppendFormat("\r\n    FROM TAXF X1 ");
            strSql.AppendFormat("\r\n   WHERE TAXNO = '{0}' ", TaxNo);
            strSql.AppendFormat("\r\n   UNION ALL ");
            strSql.AppendFormat("\r\n  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD3, 2), '-', RIGHT(MMDD3, 2)), ITNM3, SPEC3, TWGT3 ");
            strSql.AppendFormat("\r\n       , COST3, TAMT3, TVAT3, TRMK3 ");
            strSql.AppendFormat("\r\n       , TAMT3 + TVAT3 AS TOT_AMT ");
            strSql.AppendFormat("\r\n    FROM TAXF X1 ");
            strSql.AppendFormat("\r\n   WHERE TAXNO = '{0}' ", TaxNo);
            strSql.AppendFormat("\r\n   UNION ALL ");
            strSql.AppendFormat("\r\n  SELECT CONCAT(DATE_FORMAT(TDATE, '%Y'), '-', LEFT(MMDD4, 2), '-', RIGHT(MMDD4, 2)), ITNM4, SPEC4, TWGT4 ");
            strSql.AppendFormat("\r\n       , COST4, TAMT4, TVAT4, TRMK4 ");
            strSql.AppendFormat("\r\n       , TAMT4 + TVAT2 AS TOT_AMT ");
            strSql.AppendFormat("\r\n    FROM TAXF X1 ");
            strSql.AppendFormat("\r\n   WHERE TAXNO = '{0}' ", TaxNo);
            
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sTaxNo = TxtTaxNo.EditValue?.ToString();
            string sTaxGu = LkupTaxGu.EditValue?.ToString();
            string sTGubn = LkupTGubn.EditValue?.ToString();
            string sPayGb = LkupPayGb.EditValue?.ToString();
            string sCvCOd = TxtCvCod.EditValue?.ToString();
            string sTDate = DateEditTDate.EditValue?.ToString().Substring(0, 10);
            string sRK = MemoRmk.EditValue?.ToString().Trim();

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
            sMMDD1 = string.IsNullOrEmpty(sMMDD1) ? null : sMMDD1.Replace("-", "").Substring(4, 4);
            string sITNM1 = dt.Rows[0]["ITNM"]?.ToString();
            string sSPEC1 = dt.Rows[0]["SPEC"]?.ToString();
            string sTWGT1 = string.IsNullOrEmpty(dt.Rows[0]["TWGT"]?.ToString()) ? null : dt.Rows[0]["TWGT"]?.ToString();
            string sCOST1 = string.IsNullOrEmpty(dt.Rows[0]["COST"]?.ToString()) ? null : dt.Rows[0]["COST"]?.ToString();
            string sTAMT1 = string.IsNullOrEmpty(dt.Rows[0]["TAMT"]?.ToString()) ? null : dt.Rows[0]["TAMT"]?.ToString();
            string sTVAT1 = string.IsNullOrEmpty(dt.Rows[0]["TVAT"]?.ToString()) ? null : dt.Rows[0]["TVAT"]?.ToString();
            string sTRMK1 = dt.Rows[0]["TRMK"]?.ToString().Trim();

            string sMMDD2 = dt.Rows[1]["MMDD"]?.ToString();
            sMMDD2 = string.IsNullOrEmpty(sMMDD2) ? null : sMMDD2.Replace("-", "").Substring(4, 4);
            string sITNM2 = dt.Rows[1]["ITNM"]?.ToString();
            string sSPEC2 = dt.Rows[1]["SPEC"]?.ToString();
            string sTWGT2 = string.IsNullOrEmpty(dt.Rows[1]["TWGT"]?.ToString()) ? null : dt.Rows[1]["TWGT"]?.ToString();
            string sCOST2 = string.IsNullOrEmpty(dt.Rows[1]["COST"]?.ToString()) ? null : dt.Rows[1]["COST"]?.ToString();
            string sTAMT2 = string.IsNullOrEmpty(dt.Rows[1]["TAMT"]?.ToString()) ? null : dt.Rows[1]["TAMT"]?.ToString();
            string sTVAT2 = string.IsNullOrEmpty(dt.Rows[1]["TVAT"]?.ToString()) ? null : dt.Rows[1]["TVAT"]?.ToString();
            string sTRMK2 = dt.Rows[1]["TRMK"]?.ToString().Trim();

            string sMMDD3 = dt.Rows[2]["MMDD"]?.ToString();
            sMMDD3 = string.IsNullOrEmpty(sMMDD3) ? null : sMMDD3.Replace("-", "").Substring(4, 4);
            string sITNM3 = dt.Rows[2]["ITNM"]?.ToString();
            string sSPEC3 = dt.Rows[2]["SPEC"]?.ToString();
            string sTWGT3 = string.IsNullOrEmpty(dt.Rows[2]["TWGT"]?.ToString()) ? null : dt.Rows[2]["TWGT"]?.ToString();
            string sCOST3 = string.IsNullOrEmpty(dt.Rows[2]["COST"]?.ToString()) ? null : dt.Rows[2]["COST"]?.ToString();
            string sTAMT3 = string.IsNullOrEmpty(dt.Rows[2]["TAMT"]?.ToString()) ? null : dt.Rows[2]["TAMT"]?.ToString();
            string sTVAT3 = string.IsNullOrEmpty(dt.Rows[2]["TVAT"]?.ToString()) ? null : dt.Rows[2]["TVAT"]?.ToString();
            string sTRMK3 = dt.Rows[2]["TRMK"]?.ToString().Trim();

            string sMMDD4 = dt.Rows[3]["MMDD"]?.ToString();
            sMMDD4 = string.IsNullOrEmpty(sMMDD4) ? null : sMMDD4.Replace("-", "").Substring(4, 4);
            string sITNM4 = dt.Rows[3]["ITNM"]?.ToString();
            string sSPEC4 = dt.Rows[3]["SPEC"]?.ToString();
            string sTWGT4 = string.IsNullOrEmpty(dt.Rows[3]["TWGT"]?.ToString()) ? null : dt.Rows[3]["TWGT"]?.ToString();
            string sCOST4 = string.IsNullOrEmpty(dt.Rows[3]["COST"]?.ToString()) ? null : dt.Rows[3]["COST"]?.ToString();
            string sTAMT4 = string.IsNullOrEmpty(dt.Rows[3]["TAMT"]?.ToString()) ? null : dt.Rows[3]["TAMT"]?.ToString();
            string sTVAT4 = string.IsNullOrEmpty(dt.Rows[3]["TVAT"]?.ToString()) ? null : dt.Rows[3]["TVAT"]?.ToString();
            string sTRMK4 = dt.Rows[3]["TRMK"]?.ToString().Trim();
            
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
                    strSql.AppendLine(" SELECT CONCAT('TX', DATE_FORMAT(NOW(), '%Y%m%d') , LPAD(IFNULL(MAX(RIGHT(TAXNO, 5)), 0) + 1, 5, '0')) AS MAX_VAL ");
                    strSql.AppendLine("   FROM TAXF A  ");
                    strSql.AppendLine("  WHERE A.TDATE = @TDATE ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@TDATE", DateTime.Today.ToString("yyyy-MM-dd"));
                    sTaxNo = cmd.ExecuteScalar()?.ToString();
                }

                string[] sArrDate = GetDealTerm(GridViewRetr);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("TAXNO", sTaxNo);
                dicParams.Add("PDAT1", sArrDate[0]);
                dicParams.Add("PDAT2", sArrDate[1]);
                dicParams.Add("TDATE", sTDate);
                dicParams.Add("CVCOD", sCvCOd);
                dicParams.Add("TGUBN", sTGubn);
                dicParams.Add("PAYGB", sPayGb);
                dicParams.Add("TAXGU", sTaxGu);
                dicParams.Add("AUTGB", "2");
                dicParams.Add("AAUTO", "B01"); //AC
                dicParams.Add("RK", sRK);

                dicParams.Add("MMDD1", sMMDD1);
                dicParams.Add("ITNM1", sITNM1);
                dicParams.Add("SPEC1", sSPEC1);
                dicParams.Add("TWGT1", sTWGT1);
                dicParams.Add("COST1", sCOST1);
                dicParams.Add("TAMT1", sTAMT1);
                dicParams.Add("TVAT1", sTVAT1);
                dicParams.Add("TRMK1", sTRMK1);

                dicParams.Add("MMDD2", sMMDD2);
                dicParams.Add("ITNM2", sITNM2);
                dicParams.Add("SPEC2", sSPEC2);
                dicParams.Add("TWGT2", sTWGT2);
                dicParams.Add("COST2", sCOST2);
                dicParams.Add("TAMT2", sTAMT2);
                dicParams.Add("TVAT2", sTVAT2);
                dicParams.Add("TRMK2", sTRMK2);

                dicParams.Add("MMDD3", sMMDD3);
                dicParams.Add("ITNM3", sITNM3);
                dicParams.Add("SPEC3", sSPEC3);
                dicParams.Add("TWGT3", sTWGT3);
                dicParams.Add("COST3", sCOST3);
                dicParams.Add("TAMT3", sTAMT3);
                dicParams.Add("TVAT3", sTVAT3);
                dicParams.Add("TRMK3", sTRMK3);

                dicParams.Add("MMDD4", sMMDD4);
                dicParams.Add("ITNM4", sITNM4);
                dicParams.Add("SPEC4", sSPEC4);
                dicParams.Add("TWGT4", sTWGT4);
                dicParams.Add("COST4", sCOST4);
                dicParams.Add("TAMT4", sTAMT4);
                dicParams.Add("TVAT4", sTVAT4);
                dicParams.Add("TRMK4", sTRMK4);

                dicParams.Add("CUSER", FmMainToolBar2.UserID);
                dicParams.Add("MUSER", FmMainToolBar2.UserID);

                strSql.Clear();
                strSql.AppendLine(" INSERT INTO TAXF ");
                strSql.AppendLine("           ( TAXNO, TDATE, PDAT1, PDAT2, CVCOD ");
                strSql.AppendLine("           , TGUBN, PAYGB, TAXGU, AUTGB, RK, AAUTO ");
                strSql.AppendLine("           , MMDD1, ITNM1, SPEC1, TWGT1, COST1, TAMT1, TVAT1, TRMK1 ");
                strSql.AppendLine("           , MMDD2, ITNM2, SPEC2, TWGT2, COST2, TAMT2, TVAT2, TRMK2 ");
                strSql.AppendLine("           , MMDD3, ITNM3, SPEC3, TWGT3, COST3, TAMT3, TVAT3, TRMK3 ");
                strSql.AppendLine("           , MMDD4, ITNM4, SPEC4, TWGT4, COST4, TAMT4, TVAT4, TRMK4 ");
                strSql.AppendLine("           , CDATE, CUSER ) ");
                strSql.AppendLine("     VALUES( @TAXNO, @TDATE, @PDAT1, @PDAT2, @CVCOD ");
                strSql.AppendLine("           , @TGUBN, @PAYGB, @TAXGU, @AUTGB, @RK, @AAUTO");
                strSql.AppendLine("           , @MMDD1, @ITNM1, @SPEC1, @TWGT1, @COST1, @TAMT1, @TVAT1, @TRMK1 ");
                strSql.AppendLine("           , @MMDD2, @ITNM2, @SPEC2, @TWGT2, @COST2, @TAMT2, @TVAT2, @TRMK2 ");
                strSql.AppendLine("           , @MMDD3, @ITNM3, @SPEC3, @TWGT3, @COST3, @TAMT3, @TVAT3, @TRMK3 ");
                strSql.AppendLine("           , @MMDD4, @ITNM4, @SPEC4, @TWGT4, @COST4, @TAMT4, @TVAT4, @TRMK4 ");
                strSql.AppendLine("           , NOW(), @CUSER ) ");
                strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                strSql.AppendLine("             TAXNO = @TAXNO, TDATE = @TDATE, PDAT1 = @PDAT1, PDAT2 = @PDAT2, CVCOD = @CVCOD ");
                strSql.AppendLine("           , TGUBN = @TGUBN, PAYGB = @PAYGB, TAXGU = @TAXGU, AUTGB = @AUTGB, RK = @RK ");
                strSql.AppendLine("           , MMDD1 = @MMDD1, ITNM1 = @ITNM1, SPEC1 = @SPEC1, TWGT1 = @TWGT1, COST1 = @COST1, TAMT1 = @TAMT1, TVAT1 = @TVAT1, TRMK1 = @TRMK1 ");
                strSql.AppendLine("           , MMDD2 = @MMDD2, ITNM2 = @ITNM2, SPEC2 = @SPEC2, TWGT2 = @TWGT2, COST2 = @COST2, TAMT2 = @TAMT2, TVAT2 = @TVAT2, TRMK2 = @TRMK2 ");
                strSql.AppendLine("           , MMDD3 = @MMDD3, ITNM3 = @ITNM3, SPEC3 = @SPEC3, TWGT3 = @TWGT3, COST3 = @COST3, TAMT3 = @TAMT3, TVAT3 = @TVAT3, TRMK3 = @TRMK3 ");
                strSql.AppendLine("           , MMDD4 = @MMDD4, ITNM4 = @ITNM4, SPEC4 = @SPEC4, TWGT4 = @TWGT4, COST4 = @COST4, TAMT4 = @TAMT4, TVAT4 = @TVAT4, TRMK4 = @TRMK4 ");
                strSql.AppendLine("           , CDATE = NOW(), CUSER = @CUSER ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                foreach (KeyValuePair<string, string> param in dicParams)
                {
                    cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                }
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                
                strSql.Clear();
                strSql.AppendFormat(" ");
                strSql.AppendFormat(" UPDATE ACTRAN ");
                strSql.AppendFormat("    SET TAXNO = '{0}' ", sTaxNo);
                strSql.AppendFormat("  WHERE TDATE = '{0}' ", _TDATE);
                strSql.AppendFormat("    AND ATGUB = '{0}' ", _ATGUB);
                strSql.AppendFormat("    AND SEQNO = '{0}' ", _SEQNO);

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                Dispose();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
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
                        view.FocusedColumn = GridColTaxMmDd;
                        bYn = false;
                    }
                    else if (string.IsNullOrEmpty(sITNM))
                    {
                        XtraMessageBox.Show("품명을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxItCd;
                        bYn = false;
                    }
                    else if (dTWGT <= 0)
                    {
                        XtraMessageBox.Show("수량을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxQty ;
                        bYn = false;
                    }
                    else if (dCOST <= 0)
                    {
                        XtraMessageBox.Show("단가를 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxUnitPrc;
                        bYn = false;
                    }
                }
                else if (string.IsNullOrEmpty(sMMDD) || string.IsNullOrEmpty(sITNM) || dTWGT == 0 || dCOST == 0)
                {
                    if (string.IsNullOrEmpty(sMMDD))
                    {
                        XtraMessageBox.Show("월일을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxMmDd;
                        bYn = false;
                    }
                    else if (string.IsNullOrEmpty(sITNM))
                    {
                        XtraMessageBox.Show("품명을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxItCd;
                        bYn = false;
                    }
                    else if (dTWGT <= 0)
                    {
                        XtraMessageBox.Show("수량을 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxQty;
                        bYn = false;
                    }
                    else if (dCOST <= 0)
                    {
                        XtraMessageBox.Show("단가를 입력하세요.");
                        view.FocusedRowHandle = i;
                        view.FocusedColumn = GridColTaxUnitPrc;
                        bYn = false;
                    }
                }
            }

            if (view.RowCount == iNotInputCnt)
            {
                XtraMessageBox.Show("상세내역을 한건 이상 입력하여야 합니다.");
                bYn = false;
            }

            return bYn;
        }

        private string[] GetDealTerm(GridView view)
        {
            string[] sArr = new string[2];

            Dictionary<int, DateTime> dicTemp = new Dictionary<int, DateTime>();
            for (int i = 0; i < view.RowCount; i++)
            {
                string sMMDD = view.GetRowCellValue(i, GridColTaxMmDd)?.ToString();
                if (string.IsNullOrEmpty(sMMDD))
                    continue;

                dicTemp.Add(i, Convert.ToDateTime(sMMDD));
            }

            sArr[0] = dicTemp.Values.Min().ToString("yyyy-MM-dd");
            sArr[1] = dicTemp.Values.Max().ToString("yyyy-MM-dd");

            return sArr;
        }

        private void AC02001F08_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewSlip_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewSlip_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }
    }
}