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
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class AC02001F04 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F04()
        {
            InitializeComponent();
        }

        public DataRow DrSlipInfo;

        private void AC02001F04_Load(object sender, EventArgs e)
        {
            DataTable dtBKind = GetLookUpData("1", "Y", "Y");
            DataTable dtDealer = GetLookUpData("2", "Y", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupBKind, dtBKind, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupDealer, dtDealer, "CD", "NM", "Y");
            
            string sTDate = DrSlipInfo["TDATE"]?.ToString();
            string sAtGub = DrSlipInfo["ATGUB"]?.ToString();
            string sSeqNo = DrSlipInfo["SEQNO"]?.ToString();
            string sLinNo = DrSlipInfo["LINNO"]?.ToString();

            DataTable dt = GetSlipRemarkInfo(sTDate.Replace("-", "").Substring(0, 8), sAtGub, sSeqNo, sLinNo);

            if (dt.Rows.Count > 0)
            {
                DateEditSlip.EditValue = dt.Rows[0]["TDATE"];
                TxtDealerCd.EditValue = dt.Rows[0]["CVCOD"];
                TxtDealerNm.EditValue = dt.Rows[0]["CVNAM"];
                TxtJunkyo.EditValue = dt.Rows[0]["ATEXT"];
                TxtBillNo.EditValue = dt.Rows[0]["BILNO"];
                LkupBKind.EditValue = dt.Rows[0]["BKIND"];
                DateEditBillFrom.EditValue = dt.Rows[0]["BSDAT"];
                DateEditBillTo.EditValue = dt.Rows[0]["BEDAT"];
                LkupDealer.EditValue = dt.Rows[0]["BBALH"];
                BtnEditBankCd.EditValue = dt.Rows[0]["BBKCD"];
                TxtBankNm.EditValue = dt.Rows[0]["COM_NM"];
                RdgbBsStat.EditValue = dt.Rows[0]["BSTAT"];
                TxtBigo.EditValue = dt.Rows[0]["DBIGO"];
            }
        }
        
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

            if (sGb.Equals("1")) //어음종류
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_03'");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT CAST(DEALER_CD AS VARCHAR) AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE A.EOB_YN = 'N' ");
                strSql.AppendLine("    AND (A.DEALER_GB = '매입' OR A.DEALER_GB = '매출' OR A.DEALER_GB = '입출') ");
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

        #region[Execute By Query]

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sTDate = DrSlipInfo["TDATE"]?.ToString().Replace("-", "").Substring(0, 8);
                string sAtGub = DrSlipInfo["ATGUB"]?.ToString();
                string sSeqNo = DrSlipInfo["SEQNO"]?.ToString();
                string sLinNo = DrSlipInfo["LINNO"]?.ToString();

                string sBillNo = TxtBillNo.EditValue?.ToString();
                string sBKind = LkupBKind.EditValue?.ToString();
                string sBillFrom = DateEditBillFrom.EditValue?.ToString().Substring(0, 10);
                string sBillTo = DateEditBillTo.EditValue?.ToString().Substring(0, 10);
                string sBbaLh = LkupDealer.EditValue?.ToString();
                string sBbkCd = BtnEditBankCd.EditValue?.ToString();
                string sBsTat = RdgbBsStat.EditValue?.ToString();
                string sBigo = TxtBigo.EditValue?.ToString();
                string sId = FmMainToolBar2.UserID;

                strSql.Clear();
                strSql.AppendLine(" ");

                #region mariaDB
                //strSql.AppendLine(" INSERT INTO ACBILL ");
                //strSql.AppendLine("           ( TDATE ");
                //strSql.AppendLine("           , ATGUB ");
                //strSql.AppendLine("           , SEQNO ");
                //strSql.AppendLine("           , LINNO ");
                //strSql.AppendLine("           , BILNO ");
                //strSql.AppendLine("           , BSTAT ");
                //strSql.AppendLine("           , BKIND ");
                //strSql.AppendLine("           , BBALH ");
                //strSql.AppendLine("           , BSDAT ");
                //strSql.AppendLine("           , BEDAT ");
                //strSql.AppendLine("           , BBKCD ");
                //strSql.AppendLine("           , DBIGO ");
                //strSql.AppendLine("           , CUSER ");
                //strSql.AppendLine("           , CDATE ) ");
                //strSql.AppendLine("     VALUES( '" + sTDate + "' ");
                //strSql.AppendLine("           , '" + sAtGub + "' ");
                //strSql.AppendLine("           , '" + sSeqNo + "' ");
                //strSql.AppendLine("           ,  " + sLinNo + " ");
                //strSql.AppendLine("           , '" + sBillNo + "' ");
                //strSql.AppendLine("           , '" + sBsTat + "' ");
                //strSql.AppendLine("           , '" + sBKind + "' ");
                //strSql.AppendLine("           , '" + sBbaLh + "' ");
                //strSql.AppendLine("           , '" + sBillFrom + "' ");
                //strSql.AppendLine("           , '" + sBillTo + "' ");
                //strSql.AppendLine("           , '" + sBbkCd + "' ");
                //strSql.AppendLine("           , '" + sBigo + "' ");
                //strSql.AppendLine("           , '" + sId + "' ");
                //strSql.AppendLine("           , CURRENT_TIMESTAMP() ) ");
                //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                //strSql.AppendLine("             BILNO = '" + sBillNo + "' ");
                //strSql.AppendLine("           , BSTAT = '" + sBsTat + "'  ");
                //strSql.AppendLine("           , BKIND = '" + sBKind + "'  ");
                //strSql.AppendLine("           , BBALH = '" + sBbaLh + "' ");
                //strSql.AppendLine("           , BSDAT = '" + sBillFrom + "' ");
                //strSql.AppendLine("           , BEDAT = '" + sBillTo + "' ");
                //strSql.AppendLine("           , BBKCD = '" + sBbkCd + "' ");
                //strSql.AppendLine("           , DBIGO = '" + sBigo + "' ");
                //strSql.AppendLine("           , MUSER = '" + sId + "' ");
                //strSql.AppendLine("           , MDATE = CURRENT_TIMESTAMP() ");
                #endregion

                strSql.AppendLine("IF EXISTS(SELECT* FROM ACBILL WHERE TDATE = '"+ sTDate + "' AND ATGUB = '"+ sAtGub + "' AND SEQNO = '"+ sSeqNo + "' AND LINNO = "+ sLinNo + ")");
                strSql.AppendLine("   BEGIN                                                                                   ");
                strSql.AppendLine("         UPDATE ACBILL                                                                     ");
                strSql.AppendLine("            SET BILNO = '" + sBillNo + "' ");
	            strSql.AppendLine("              , BSTAT = '" + sBsTat + "'  ");
	            strSql.AppendLine("              , BKIND = '" + sBKind + "'  ");
	            strSql.AppendLine("              , BBALH = '" + sBbaLh + "' ");
	            strSql.AppendLine("              , BSDAT = '" + sBillFrom + "' ");
	            strSql.AppendLine("              , BEDAT = '" + sBillTo + "' ");
	            strSql.AppendLine("              , BBKCD = '" + sBbkCd + "' ");
	            strSql.AppendLine("              , DBIGO = '" + sBigo + "' ");
	            strSql.AppendLine("              , MUSER = '" + sId + "' ");
	            strSql.AppendLine("              , MDATE = CONVERT(VARCHAR(19),GETDATE(),20)");
                strSql.AppendLine("          WHERE TDATE = '" + sTDate + "' AND ATGUB = '" + sAtGub + "' AND SEQNO = '" + sSeqNo + "' AND LINNO = " + sLinNo);
                strSql.AppendLine("     END                                                              ");
                strSql.AppendLine("ELSE                                                                  ");
                strSql.AppendLine("   BEGIN");
                strSql.AppendLine("         INSERT INTO ACBILL ");
                strSql.AppendLine("              (TDATE ");
                strSql.AppendLine("              , ATGUB ");
                strSql.AppendLine("              , SEQNO ");
                strSql.AppendLine("              , LINNO ");
                strSql.AppendLine("              , BILNO ");
                strSql.AppendLine("              , BSTAT ");
                strSql.AppendLine("              , BKIND ");
                strSql.AppendLine("              , BBALH ");
                strSql.AppendLine("              , BSDAT ");
                strSql.AppendLine("              , BEDAT ");
                strSql.AppendLine("              , BBKCD ");
                strSql.AppendLine("              , DBIGO ");
                strSql.AppendLine("              , CUSER ");
                strSql.AppendLine("              , CDATE) ");
                strSql.AppendLine("        VALUES('" + sTDate + "' ");
                strSql.AppendLine("              , '" + sAtGub + "' ");
                strSql.AppendLine("              , '" + sSeqNo + "' ");
                strSql.AppendLine("              , " + sLinNo + " ");
                strSql.AppendLine("              , '" + sBillNo + "' ");
                strSql.AppendLine("              , '" + sBsTat + "' ");
                strSql.AppendLine("              , '" + sBKind + "' ");
                strSql.AppendLine("              , '" + sBbaLh + "' ");
                strSql.AppendLine("              , '" + sBillFrom + "' ");
                strSql.AppendLine("              , '" + sBillTo + "' ");
                strSql.AppendLine("              , '" + sBbkCd + "' ");
                strSql.AppendLine("              , '" + sBigo + "' ");
                strSql.AppendLine("              , '" + sId + "' ");
                strSql.AppendLine("              , CONVERT(VARCHAR(19),GETDATE(),20)) ");
                strSql.AppendLine("     END");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private DataTable GetSlipRemarkInfo(string sTDate, string sAtGub, string sSeqNo, string sLinNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT CONVERT(DATE,A.TDATE) AS TDATE ");
            strSql.AppendLine(" 	 , A.CVCOD ");
            strSql.AppendLine(" 	 , B.DEALER_NM AS CVNAM ");
            strSql.AppendLine(" 	 , A.ATEXT  ");
            strSql.AppendLine("      , C.BILNO ");
            strSql.AppendLine(" 	 , C.BSTAT  ");
            strSql.AppendLine(" 	 , C.BKIND  ");
            strSql.AppendLine(" 	 , C.BBALH  ");
            strSql.AppendLine(" 	 , C.BSDAT  ");
            strSql.AppendLine(" 	 , C.BEDAT  ");
            strSql.AppendLine(" 	 , C.BBKCD  ");
            strSql.AppendLine("      , D.COM_NM ");
            strSql.AppendLine(" 	 , C.DBIGO  ");
            strSql.AppendLine("   FROM ACTRAN A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B ");
            strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACBILL C ");
            strSql.AppendLine("     ON A.TDATE = C.TDATE ");
            strSql.AppendLine("    AND A.ATGUB = C.ATGUB  ");
            strSql.AppendLine("    AND A.SEQNO = C.SEQNO  ");
            strSql.AppendLine("    AND A.LINNO = C.LINNO  ");
            strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD D ");
            strSql.AppendLine("     ON C.BBKCD = D.COM_CD ");
            strSql.AppendLine("    AND D.CD_GB = 'BANK_CD' ");
            strSql.AppendLine("  WHERE A.TDATE = '" + sTDate + "' ");
            strSql.AppendLine("    AND A.ATGUB = '" + sAtGub + "' ");
            strSql.AppendLine("    AND A.SEQNO = '" + sSeqNo + "' ");
            strSql.AppendLine("    AND A.LINNO = " + sLinNo + " ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #endregion[Execute By Query]
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void AC02001F04_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }

        public DataRow DrBankInfo;
        private void BtnEditBankCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            AC02001F05 frm = new AC02001F05();
            frm.BankCd = BtnEditBankCd.EditValue?.ToString();
            frm.P_AC02001F04 = this;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                if(DrBankInfo != null)
                {
                    BtnEditBankCd.EditValue = DrBankInfo["BANK_CD"];
                    TxtBankNm.EditValue = DrBankInfo["BANK_NM"];
                }
            }
        }
    }
}