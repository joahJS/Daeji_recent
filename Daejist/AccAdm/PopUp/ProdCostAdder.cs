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
    public partial class ProdCostAdder : DevExpress.XtraEditors.XtraForm
    {
        public ProdCostAdder()
        {
            InitializeComponent();
        }

        public string sMakeNo { get; set; }
        public string sMakeNoLn { get; set; }
        public string sEmpID { get; set; }
        public string ProcessDate { get; set; }
        public string sEmpNm { get; set; }

        private void ProdCostAdder_Load(object sender, EventArgs e)
        {
            DataTable dtDealer = GetLookUpData("2", "Y", "Y");
            DataTable dtEquip = GetLookUpData("3", "Y", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupDealerNm, dtDealer, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupDeviceNm, dtEquip, "CD", "NM", "Y");

            if (!string.IsNullOrEmpty(ProcessDate))
            {
                if (ProcessDate.Length == 8)
                {
                    string sResult = ProcessDate.Substring(0, 4) + "-" + ProcessDate.Substring(4, 2) + "-" + ProcessDate.Substring(6, 2);
                    DateEditYmd.EditValue = sResult;
                }
            }

            GetMakeExpenseInfo();
        }

        #region[LookupDate Setting]
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            
            if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT CONVERT(VARCHAR,A.DEALER_CD) AS CD");
                strSql.AppendLine("      , A.DEALER_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_CD) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE EOB_YN = 'N'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.MG_NO AS CD");
                strSql.AppendLine("      , A.EQUIP_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY MG_NO) AS SEQ");
                strSql.AppendLine("   FROM EQUIP_CD A ");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }
        #endregion[LookupDate Setting]

        private void GetMakeExpenseInfo()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LN ");
            strSql.AppendLine("      , A.LN_ESEQ ");
            strSql.AppendLine("      , A.GUBUN ");
            strSql.AppendLine("      , A.EDATE ");
            strSql.AppendLine("      , A.EGUBUN ");
            strSql.AppendLine("      , A.EUSER ");
            strSql.AppendLine("      , A.EEQUIP ");
            strSql.AppendLine("      , A.ECVNAM ");
            strSql.AppendLine("      , A.ECONTENT ");
            strSql.AppendLine("      , A.EREPAIR ");
            strSql.AppendLine("      , A.EKIND_L ");
            strSql.AppendLine("      , A.EKIND_M ");
            strSql.AppendLine("      , A.EITCOD ");
            strSql.AppendLine("      , A.EQTY ");
            strSql.AppendLine("      , A.EDAN ");
            strSql.AppendLine("      , A.EAMT ");
            strSql.AppendLine("      , A.WDATE ");
            strSql.AppendLine("   FROM MAKE_EXPENSE A ");
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
            strSql.AppendLine("    AND MAKENO_LN = " + sMakeNoLn + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            if(dt.Rows.Count > 0)
            {
                string sEquip = dt.Rows[0]["EEQUIP"]?.ToString();
                string sDealer = dt.Rows[0]["ECVNAM"]?.ToString();
                string sDate = dt.Rows[0]["EDATE"]?.ToString();
                string sRemark = dt.Rows[0]["ECONTENT"]?.ToString();
                string sRepair = dt.Rows[0]["EREPAIR"]?.ToString();
                string sInOutGB = dt.Rows[0]["EGUBUN"]?.ToString();
                string sBigCate = dt.Rows[0]["EKIND_L"]?.ToString();
                string sMidCate = dt.Rows[0]["EKIND_M"]?.ToString();

                if (!string.IsNullOrEmpty(sDate) && sDate.Length == 8)
                {
                    DateEditYmd.EditValue = sDate.Substring(0, 4) + "-" + sDate.Substring(4, 2) + "-" + sDate.Substring(6, 2);
                }
                LkupDeviceNm.Text = sEquip;
                LkupDealerNm.Text = sDealer;
                TxtBreakResn.Text = sRemark;
                TxtRepair.Text = sRepair;
                RdgbInOutGb.EditValue = sInOutGB;
                RdgbBigCate.EditValue = sBigCate;
                RdgbMidCate.EditValue = sMidCate;
            }
            else
            {
                Cursor = Cursors.WaitCursor;

                DataTable dtChk = (DataTable)GridRetr.DataSource;
                DataRow row = dtChk.NewRow();
                row["MAKENO"] = sMakeNo;
                row["MAKENO_LN"] = sMakeNoLn;
                row["LN_ESEQ"] = dtChk.Rows.Count;

                dtChk.Rows.Add(row);
                GridRetr.DataSource = dtChk;
                //GridViewRetr.AddNewRow();
                //GridViewRetr.SetFocusedRowCellValue("MAKENO", sMakeNo);
                //GridViewRetr.SetFocusedRowCellValue("MAKENO_LN", sMakeNoLn);
                //GridViewRetr.SetFocusedRowCellValue("LN_ESEQ", GridViewRetr.RowCount - 1);

                Cursor = Cursors.Default;
            }
        }

        private bool MakeNoApprovalYN(string MDate)
        {
            bool Aprv = false;
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT CASE WHEN SIGN4 = 'Y' THEN 'Y' ELSE 'N' END AS APRV ");
            strSql.AppendLine("   FROM MAKE_S ");
            strSql.AppendLine("  WHERE MDATE = '" + MDate + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count > 0)
            {
                string sAprvYN = dt.Rows[0]["APRV"]?.ToString();
                if (string.IsNullOrEmpty(sAprvYN))
                {
                    Aprv = false;
                }
                else if (sAprvYN.Equals("Y"))
                {
                    Aprv = true;
                }
                else if(sAprvYN.Equals("N"))    
                {
                    Aprv = false;
                }
            }
            else
            {
                Aprv = false;
            }

            return Aprv;
        }

        private void BtnSaveClose_Click(object sender, EventArgs e)
        {
            string sDateChk = DateEditYmd.EditValue?.ToString();
            string sDvcNm = LkupDeviceNm.Text;
            string sDlrNm = LkupDealerNm.Text;

            string sDate = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            if (MakeNoApprovalYN(sDate))
            {
                XtraMessageBox.Show("해당 데이터는 결제완료 건입니다.");
                return;
            }


            //전표 발행여부 체크
            StringBuilder strSql1 = new StringBuilder();

            strSql1.Clear();
            strSql1.AppendLine(" SELECT CASE WHEN A.SLIP_YN IS NULL THEN 'N' ELSE A.SLIP_YN END AS SLIP_YN ");
            strSql1.AppendLine("   FROM EQUIP_CD_HISTORY A ");
            strSql1.AppendLine("   LEFT OUTER JOIN MAKE_EXPENSE B ");
            strSql1.AppendLine("     ON A.MAKENO = B.MAKENO ");
            strSql1.AppendLine("    AND A.MAKENO_LN = B.MAKENO_LN ");
            strSql1.AppendLine("  WHERE A.MAKENO = " + sMakeNo + " ");
            strSql1.AppendLine("    AND A.MAKENO_LN = " + sMakeNoLn + " ");

            DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, strSql1.ToString());
            if(dt1.Rows.Count > 0)
            {
                string sSlipYn = dt1.Rows[0]["SLIP_YN"]?.ToString();
                if (sSlipYn.Equals("Y"))
                {
                    XtraMessageBox.Show("해당 비용 건은 전표승인 상태입니다.");
                    return;
                }
            }

            if (string.IsNullOrEmpty(sDateChk))
            {
                XtraMessageBox.Show("발생일자를 선택하여 주세요.");
                return;
            }
            else if (string.IsNullOrEmpty(sDvcNm))
            {
                XtraMessageBox.Show("장비를 선택하여 주세요.");
                return;
            }
            else if (string.IsNullOrEmpty(sDlrNm))
            {
                XtraMessageBox.Show("거래처를 선택하여 주세요.");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                Cursor = Cursors.Default;

                StringBuilder strSql = new StringBuilder();

                string sUserEmpID = string.Empty;
                string sPrcsYmd = string.Empty;
                string sDeviceNm = string.Empty;
                string sDeviceCd = string.Empty;
                string sBrokenResn = string.Empty;
                string sRepair = string.Empty;
                string sDealerNm = string.Empty;
                string sBigCate = string.Empty;
                string sMidCate = string.Empty;
                string sInOutGB = string.Empty;

                sUserEmpID = sEmpID;
                sPrcsYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
                sDeviceNm = LkupDeviceNm.Text;
                sDeviceCd = LkupDeviceNm.EditValue?.ToString();
                sDealerNm = LkupDealerNm.Text;
                sBrokenResn = TxtBreakResn.EditValue?.ToString();
                sRepair = TxtRepair.EditValue?.ToString();
                sBigCate = RdgbBigCate.EditValue?.ToString();
                sMidCate = RdgbMidCate.EditValue?.ToString();
                sInOutGB = RdgbInOutGb.EditValue?.ToString();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.MDATE ");
                strSql.AppendLine("   FROM MAKE_M A ");
                strSql.AppendLine("  WHERE MAKENO = " + sMakeNo +"");
                strSql.AppendLine("    AND MUSER_ID = " + sEmpID + "");

                DataTable dtMakeM = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dtMakeM.Rows.Count == 0)
                { 
                    #region[MAKE_M 등록 부분]

                    string sEntID = FmMainToolBar2.UserID;

                    string sWorkingYmd = DateEditYmd.EditValue.ToString().Replace("-", "").Substring(0, 8);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT A.MDATE ");
                    strSql.AppendLine("   FROM MAKE_S A ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkingYmd + "' ");

                    DataTable dtMakeS = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    if (dtMakeS.Rows.Count == 0)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" INSERT INTO MAKE_S ");
                        strSql.AppendLine("           ( MDATE ");
                        strSql.AppendLine("           , GUBUN ");
                        strSql.AppendLine("           , SIGN1 ");
                        strSql.AppendLine("           , SIGN2 ");
                        strSql.AppendLine("           , SIGN3 ");
                        strSql.AppendLine("           , MCLOSED ");
                        strSql.AppendLine("           , MLATENESS ");
                        strSql.AppendLine("           , MLEAVE ");
                        strSql.AppendLine("           , MGOOUT ");
                        strSql.AppendLine("           ) ");
                        strSql.AppendLine("      VALUES ");
                        strSql.AppendLine("           ( '" + sWorkingYmd + "' ");
                        strSql.AppendLine("           , '1' ");
                        strSql.AppendLine("           , 'N' ");
                        strSql.AppendLine("           , 'N'");
                        strSql.AppendLine("           , 'N'");
                        strSql.AppendLine("           , 0 ");
                        strSql.AppendLine("           , 0 ");
                        strSql.AppendLine("           , 0 ");
                        strSql.AppendLine("           , 0 ");
                        strSql.AppendLine("           ) ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO MAKE_M ");
                    //strSql.AppendLine("           ( ");
                    //strSql.AppendLine("             MAKENO ");
                    //strSql.AppendLine("           , MDATE ");
                    //strSql.AppendLine("           , MUSER ");
                    //strSql.AppendLine("           , MUSER_ID ");
                    //strSql.AppendLine("           , ENT_DT ");
                    //strSql.AppendLine("           , ENT_ID ");
                    //strSql.AppendLine("           , MFY_DT ");
                    //strSql.AppendLine("           , MFY_ID ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("      VALUES ");
                    //strSql.AppendLine("           ( ");
                    //strSql.AppendLine("             " + sMakeNo + " ");
                    //strSql.AppendLine("           , '" + sWorkingYmd + "' ");
                    //strSql.AppendLine("           , '" + sEmpNm + "' ");
                    //strSql.AppendLine("           , '" + sEmpID + "' ");
                    //strSql.AppendLine("           , NOW() ");
                    //strSql.AppendLine("           , '" + sEntID + "' ");
                    //strSql.AppendLine("           , NOW()");
                    //strSql.AppendLine("           , '" + sEntID + "' ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             MDATE = '" + sWorkingYmd + "' ");
                    //strSql.AppendLine("           , MUSER = '" + sEmpNm + "' ");
                    //strSql.AppendLine("           , MUSER_ID = '" + sEmpID + "' ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID = '" + sEntID + "' ");
                    #endregion
                    strSql.AppendLine("IF EXISTS(SELECT FROM MAKE_M WHERE MAKENO = " + sMakeNo + ")");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         UPDATE MAKE_M");
                    strSql.AppendLine("            SET MDATE = '" + sWorkingYmd + "' ");
                    strSql.AppendLine("              , MUSER = '" + sEmpNm + "' ");
                    strSql.AppendLine("              , MUSER_ID = '" + sEmpID + "' ");
                    strSql.AppendLine("              , MFY_DT = CONVERT([varchar](20), getdate(), (21)) ");
                    strSql.AppendLine("              , MFY_ID = '" + sEntID + "' ");
                    strSql.AppendLine("          WHERE MAKENO = " + sMakeNo);
                    strSql.AppendLine("   END");
                    strSql.AppendLine("ELSE");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         INSERT INTO MAKE_M ");
                    strSql.AppendLine("                   ( ");
                    strSql.AppendLine("                     MAKENO ");
                    strSql.AppendLine("                   , MDATE ");
                    strSql.AppendLine("                   , MUSER ");
                    strSql.AppendLine("                   , MUSER_ID ");
                    strSql.AppendLine("                   , ENT_DT ");
                    strSql.AppendLine("                   , ENT_ID ");
                    strSql.AppendLine("                   , MFY_DT ");
                    strSql.AppendLine("                   , MFY_ID ");
                    strSql.AppendLine("                   ) ");
                    strSql.AppendLine("              VALUES ");
                    strSql.AppendLine("                   ( ");
                    strSql.AppendLine("                     " + sMakeNo + " ");
                    strSql.AppendLine("                   , '" + sWorkingYmd + "' ");
                    strSql.AppendLine("                   , '" + sEmpNm + "' ");
                    strSql.AppendLine("                   , '" + sEmpID + "' ");
                    strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21)) ");
                    strSql.AppendLine("                   , '" + sEntID + "' ");
                    strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21))");
                    strSql.AppendLine("                   , '" + sEntID + "' ");
                    strSql.AppendLine("                   ) ");
                    strSql.AppendLine("   END");

                    
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    #endregion[MAKE_M 등록 부분]
                }

                DataTable dt = (DataTable)GridRetr.DataSource;

                double dSumAmt = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sSaveMakeNo = string.Empty;
                    string sMakeNoLn = string.Empty;
                    string sMakeSeq = string.Empty;
                    string sItemNm = string.Empty;
                    string sQuantity = string.Empty;
                    string sUnitPrc = string.Empty;
                    string sTotAmt = string.Empty;
                    
                    double dQuantity = 0;
                    double dUnitPrc = 0;
                    double dTotAmt = 0;

                    sSaveMakeNo = dt.Rows[i]["MAKENO"].ToString();
                    sMakeNoLn   = dt.Rows[i]["MAKENO_LN"].ToString();
                    sMakeSeq    = dt.Rows[i]["LN_ESEQ"].ToString();
                    sItemNm     = dt.Rows[i]["EITCOD"].ToString();
                    sQuantity   = dt.Rows[i]["EQTY"].ToString();
                    sUnitPrc    = dt.Rows[i]["EDAN"].ToString();
                    sTotAmt     = dt.Rows[i]["EAMT"].ToString();

                    dQuantity = string.IsNullOrEmpty(sQuantity) ? 0 : Convert.ToDouble(sQuantity);
                    dUnitPrc = string.IsNullOrEmpty(sUnitPrc) ? 0 : Convert.ToDouble(sUnitPrc);
                    //dTotAmt = string.IsNullOrEmpty(sTotAmt) ? 0 : Convert.ToDouble(sTotAmt);
                    dTotAmt = dQuantity * dUnitPrc;
                    dSumAmt += dTotAmt;

                    if (i == 0 || i == dt.Rows.Count - 1)
                    {
                        string ChksItemNm = dt.Rows[i]["EITCOD"].ToString();
                        string ChksQuantity = dt.Rows[i]["EQTY"].ToString();
                        string ChksUnitPrc = dt.Rows[i]["EDAN"].ToString();
                        string ChksTotAmt = dt.Rows[i]["EAMT"].ToString();

                        if (string.IsNullOrEmpty(ChksItemNm) & string.IsNullOrEmpty(ChksQuantity)
                          & string.IsNullOrEmpty(ChksUnitPrc) & string.IsNullOrEmpty(ChksTotAmt))
                        {
                            continue;
                        }
                    }

                    //설비이력 추가/변경(MAKE_LN 값까지 합산비용금액)
                    if(i == dt.Rows.Count - 1)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT * ");
                        strSql.AppendLine("   FROM EQUIP_CD_HISTORY ");
                        strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");
                        strSql.AppendLine("    AND MAKENO_LN = " + sMakeNoLn + " ");

                        DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        if (dtChk.Rows.Count == 0)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" SELECT ISNULL(MAX(MG_HIS_SEQ), 0) AS MAX_VALUE   ");
                            strSql.AppendLine("   FROM EQUIP_CD_HISTORY ");
                            strSql.AppendLine("  WHERE MG_NO = '" + sDeviceCd + "' ");

                            DataTable dtMax = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            double dMaxValue = Convert.ToDouble(dtMax.Rows[0]["MAX_VALUE"]) + 1;
                            string sYmd = DateEditYmd.EditValue?.ToString().Substring(0, 10);

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" INSERT INTO EQUIP_CD_HISTORY ");
                            strSql.AppendLine(" 		  ( ");
                            strSql.AppendLine(" 		    MG_NO, OCCUR_DT, MG_DESC ");
                            strSql.AppendLine(" 		  , MG_COST, MG_USER, MAKENO ");
                            strSql.AppendLine(" 		  , MAKENO_LN, SLIP_YN, LN_ESEQ) ");
                            strSql.AppendLine("      VALUES ");
                            strSql.AppendLine("           ( ");
                            strSql.AppendLine("             '" + sDeviceCd + "', '" + sYmd + "', '고장내역 : " + sBrokenResn + ", 수리사항 : " + sRepair + "' ");
                            strSql.AppendLine("           , '" + dSumAmt + "', '" + sUserEmpID + "', " + sMakeNo + " ");
                            strSql.AppendLine("           , " + sMakeNoLn + ", 'N', "+ sMakeSeq + " ");
                            strSql.AppendLine("           ) ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            string sYmd = DateEditYmd.EditValue?.ToString().Substring(0, 10);

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" UPDATE EQUIP_CD_HISTORY ");
                            strSql.AppendLine("    SET OCCUR_DT = '" + sYmd + "' ");
                            strSql.AppendLine("      , MG_DESC = '고장내역 : " + sBrokenResn + ", 수리사항 : " + sRepair + "' ");
                            strSql.AppendLine("      , MG_COST = '" + dSumAmt + "' ");
                            strSql.AppendLine("      , MG_USER = '" + sUserEmpID + "' ");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + "  ");
                            strSql.AppendLine("    AND MAKENO_LN = " + sMakeNoLn + "  ");
                            strSql.AppendLine("    AND LN_ESEQ = " + sMakeSeq + "  ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                    }
                   
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO MAKE_EXPENSE ");
                    //strSql.AppendLine("           ( ");
                    //strSql.AppendLine("             MAKENO ");
                    //strSql.AppendLine("           , MAKENO_LN ");
                    //strSql.AppendLine("           , LN_ESEQ ");
                    //strSql.AppendLine("           , GUBUN ");
                    //strSql.AppendLine("           , EDATE ");
                    //strSql.AppendLine("           , EGUBUN ");
                    //strSql.AppendLine("           , EUSER ");
                    //strSql.AppendLine("           , EEQUIP ");
                    //strSql.AppendLine("           , ECVNAM ");
                    //strSql.AppendLine("           , ECONTENT ");
                    //strSql.AppendLine("           , EREPAIR ");
                    //strSql.AppendLine("           , EKIND_L ");
                    //strSql.AppendLine("           , EKIND_M ");
                    //strSql.AppendLine("           , EITCOD ");
                    //strSql.AppendLine("           , EQTY ");
                    //strSql.AppendLine("           , EDAN ");
                    //strSql.AppendLine("           , EAMT ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("      VALUES ");
                    //strSql.AppendLine("           ( ");
                    //strSql.AppendLine("             " + sSaveMakeNo + " ");
                    //strSql.AppendLine("           , " + sMakeNoLn + " ");
                    //strSql.AppendLine("           , " + sMakeSeq + " ");
                    //strSql.AppendLine("           , '1' ");
                    //strSql.AppendLine("           , '" + sPrcsYmd + "' ");
                    //strSql.AppendLine("           , '" + sInOutGB + "' ");
                    //strSql.AppendLine("           , '" + sUserEmpID + "' ");
                    //strSql.AppendLine("           , '" + sDeviceNm + "' ");
                    //strSql.AppendLine("           , '" + sDealerNm + "' ");
                    //strSql.AppendLine("           , '" + sBrokenResn + "' ");
                    //strSql.AppendLine("           , '" + sRepair + "' ");
                    //strSql.AppendLine("           , '" + sBigCate + "' ");
                    //strSql.AppendLine("           , '" + sMidCate + "' ");
                    //strSql.AppendLine("           , '" + sItemNm + "' ");
                    //strSql.AppendLine("           , " + dQuantity + " ");
                    //strSql.AppendLine("           , " + dUnitPrc + " ");
                    //strSql.AppendLine("           , " + dTotAmt + " ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             EDATE = '" + sPrcsYmd + "' ");
                    //strSql.AppendLine("           , EGUBUN = '" + sInOutGB + "' ");
                    //strSql.AppendLine("           , EUSER = '" + sUserEmpID + "' ");
                    //strSql.AppendLine("           , EEQUIP = '" + sDeviceNm + "' ");
                    //strSql.AppendLine("           , ECVNAM = '" + sDealerNm + "' ");
                    //strSql.AppendLine("           , ECONTENT = '" + sBrokenResn + "' ");
                    //strSql.AppendLine("           , EREPAIR = '" + sRepair + "' ");
                    //strSql.AppendLine("           , EKIND_L = '" + sBigCate + "' ");
                    //strSql.AppendLine("           , EKIND_M = '" + sMidCate + "' ");
                    //strSql.AppendLine("           , EITCOD = '" + sItemNm + "' ");
                    //strSql.AppendLine("           , EQTY = " + dQuantity + " ");
                    //strSql.AppendLine("           , EDAN = " + dUnitPrc + " ");
                    //strSql.AppendLine("           , EAMT = " + dTotAmt + " ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_EXPENSE WHERE MAKENO = " + sSaveMakeNo + " AND MAKENO_LN = " + sMakeNoLn + " AND LN_ESEQ = " + sMakeSeq + " AND GUBUN = '1')");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         UPDATE MAKE_EXPENSE");
                    strSql.AppendLine("            SET EDATE = '" + sPrcsYmd + "' ");
                    strSql.AppendLine("              , EGUBUN = '" + sInOutGB + "' ");
                    strSql.AppendLine("              , EUSER = '" + sUserEmpID + "' ");
                    strSql.AppendLine("              , EEQUIP = '" + sDeviceNm + "' ");
                    strSql.AppendLine("              , ECVNAM = '" + sDealerNm + "' ");
                    strSql.AppendLine("              , ECONTENT = '" + sBrokenResn + "' ");
                    strSql.AppendLine("              , EREPAIR = '" + sRepair + "' ");
                    strSql.AppendLine("              , EKIND_L = '" + sBigCate + "' ");
                    strSql.AppendLine("              , EKIND_M = '" + sMidCate + "' ");
                    strSql.AppendLine("              , EITCOD = '" + sItemNm + "' ");
                    strSql.AppendLine("              , EQTY = " + dQuantity + " ");
                    strSql.AppendLine("              , EDAN = " + dUnitPrc + " ");
                    strSql.AppendLine("              , EAMT = " + dTotAmt + " ");
                    strSql.AppendLine("          WHERE MAKENO = " + sSaveMakeNo + " AND MAKENO_LN = " + sMakeNoLn + " AND LN_ESEQ = " + sMakeSeq + " AND GUBUN = '1'");
                    strSql.AppendLine("   END");
                    strSql.AppendLine("ELSE");
                    strSql.AppendLine("   BEGIN");
                    strSql.AppendLine("         INSERT INTO MAKE_EXPENSE ");
                    strSql.AppendLine("                   ( ");
                    strSql.AppendLine("                     MAKENO ");
                    strSql.AppendLine("                   , MAKENO_LN ");
                    strSql.AppendLine("                   , LN_ESEQ ");
                    strSql.AppendLine("                   , GUBUN ");
                    strSql.AppendLine("                   , EDATE ");
                    strSql.AppendLine("                   , EGUBUN ");
                    strSql.AppendLine("                   , EUSER ");
                    strSql.AppendLine("                   , EEQUIP ");
                    strSql.AppendLine("                   , ECVNAM ");
                    strSql.AppendLine("                   , ECONTENT ");
                    strSql.AppendLine("                   , EREPAIR ");
                    strSql.AppendLine("                   , EKIND_L ");
                    strSql.AppendLine("                   , EKIND_M ");
                    strSql.AppendLine("                   , EITCOD ");
                    strSql.AppendLine("                   , EQTY ");
                    strSql.AppendLine("                   , EDAN ");
                    strSql.AppendLine("                   , EAMT ");
                    strSql.AppendLine("                   ) ");
                    strSql.AppendLine("              VALUES ");
                    strSql.AppendLine("                   ( ");
                    strSql.AppendLine("                     " + sSaveMakeNo + " ");
                    strSql.AppendLine("                   , " + sMakeNoLn + " ");
                    strSql.AppendLine("                   , " + sMakeSeq + " ");
                    strSql.AppendLine("                   , '1' ");
                    strSql.AppendLine("                   , '" + sPrcsYmd + "' ");
                    strSql.AppendLine("                   , '" + sInOutGB + "' ");
                    strSql.AppendLine("                   , '" + sUserEmpID + "' ");
                    strSql.AppendLine("                   , '" + sDeviceNm + "' ");
                    strSql.AppendLine("                   , '" + sDealerNm + "' ");
                    strSql.AppendLine("                   , '" + sBrokenResn + "' ");
                    strSql.AppendLine("                   , '" + sRepair + "' ");
                    strSql.AppendLine("                   , '" + sBigCate + "' ");
                    strSql.AppendLine("                   , '" + sMidCate + "' ");
                    strSql.AppendLine("                   , '" + sItemNm + "' ");
                    strSql.AppendLine("                   , " + dQuantity + " ");
                    strSql.AppendLine("                   , " + dUnitPrc + " ");
                    strSql.AppendLine("                   , " + dTotAmt + " ");
                    strSql.AppendLine("                   ) ");
                    strSql.AppendLine("   END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                this.DialogResult = DialogResult.OK;

                Dispose();

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void GridRetr_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                if (XtraMessageBox.Show("추가하시겠습니까?", "비용 항목 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    Cursor = Cursors.WaitCursor;

                    DataTable dt = (DataTable)GridRetr.DataSource;
                    DataRow row = dt.NewRow();
                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LN"] = sMakeNoLn;
                    row["LN_ESEQ"] = dt.Rows.Count;

                    dt.Rows.Add(row);
                    GridRetr.DataSource = dt;
                    //GridViewRetr.AddNewRow();
                    //GridViewRetr.SetFocusedRowCellValue("MAKENO", sMakeNo);
                    //GridViewRetr.SetFocusedRowCellValue("MAKENO_LN", sMakeNoLn);
                    //GridViewRetr.SetFocusedRowCellValue("LN_ESEQ", GridViewRetr.RowCount - 1);

                    Cursor = Cursors.Default;
                }
            }

            if (e.KeyCode == Keys.F4)
            {
                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "비용 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    if(GetDeleteSourceInfo(sMakeNo, sMakeNoLn, GridViewRetr.GetFocusedRowCellValue("LN_ESEQ").ToString()))
                    {
                        Cursor = Cursors.WaitCursor;

                        DeleteMakeExpenseInfo(sMakeNo, sMakeNoLn, GridViewRetr.GetFocusedRowCellValue("LN_ESEQ").ToString());

                        Cursor = Cursors.Default;
                    }
                    else
                    {
                        Cursor = Cursors.WaitCursor;
                        int idx = GridViewRetr.FocusedRowHandle;

                        GridViewRetr.FocusedRowHandle = 0;
                        GridViewRetr.FocusedRowHandle = idx;

                        DataTable dt = (DataTable)GridRetr.DataSource;

                        string sSeq = GridViewRetr.GetFocusedRowCellValue("LN_ESEQ")?.ToString();
                        
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string MakeNo = dt.Rows[i]["MAKENO"].ToString();
                            string MakeNoLn = dt.Rows[i]["MAKENO_LN"].ToString();
                            string LnSeq = dt.Rows[i]["LN_ESEQ"].ToString();

                            if(MakeNo.Equals(sMakeNo) && MakeNoLn.Equals(MakeNoLn) && LnSeq.Equals(sSeq))
                            {
                                dt.Rows.RemoveAt(i);
                                break;
                            }
                        }
                        GridRetr.DataSource = dt;

                        Cursor = Cursors.Default;
                    }
                }
            }

            if (e.KeyCode == Keys.Down)
            {
                if(GridViewRetr.RowCount == 0)
                {
                    DataTable dtChk = (DataTable)GridRetr.DataSource;
                    DataRow row = dtChk.NewRow();
                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LN"] = sMakeNoLn;
                    row["LN_ESEQ"] = dtChk.Rows.Count;

                    dtChk.Rows.Add(row);
                    GridRetr.DataSource = dtChk;
                    return;
                }

                DataTable dt = (DataTable)GridRetr.DataSource;

                int iRowHandle = GridViewRetr.FocusedRowHandle;
                if (iRowHandle < dt.Rows.Count - 1)
                {
                    return;
                }

                if (dt.Rows.Count > 0)
                {
                    string sChkEItCod = dt.Rows[dt.Rows.Count-1]["EITCOD"]?.ToString();
                    string sChkEAmt = dt.Rows[dt.Rows.Count-1]["EAMT"]?.ToString();
                    string sChkEQty = dt.Rows[dt.Rows.Count-1]["EQTY"]?.ToString();
                    string sChkEDan = dt.Rows[dt.Rows.Count-1]["EDAN"]?.ToString();

                    if(string.IsNullOrEmpty(sChkEItCod) & string.IsNullOrEmpty(sChkEAmt) //이전 Row 공백체크
                      & string.IsNullOrEmpty(sChkEQty) & string.IsNullOrEmpty(sChkEDan))
                    {
                        return;
                    }

                    string sEItCod = GridViewRetr.GetFocusedRowCellValue("EITCOD")?.ToString();
                    string sEAmt = GridViewRetr.GetFocusedRowCellValue("EAMT")?.ToString();

                    if (string.IsNullOrEmpty(sEItCod) || string.IsNullOrEmpty(sEAmt))
                    {
                        return;
                    }
                    else
                    {
                        DataRow row = dt.NewRow();
                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LN"] = sMakeNoLn;
                        row["LN_ESEQ"] = dt.Rows.Count;

                        dt.Rows.Add(row);
                        GridRetr.DataSource = dt;
                    }
                }
            }

        }
        
        private bool GetDeleteSourceInfo(string MakeNo, string MakeNoLn, string Seq)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LN ");
            strSql.AppendLine("      , A.LN_ESEQ ");
            strSql.AppendLine("   FROM MAKE_EXPENSE A ");
            strSql.AppendLine("  WHERE MAKENO = " + MakeNo + " ");
            strSql.AppendLine("    AND MAKENO_LN = " + MakeNoLn + " ");
            strSql.AppendLine("    AND LN_ESEQ = " + Seq + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt.Rows.Count > 0)
            {
                return true;
            }
            else if(dt.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                return false;
            }
        }

        private void DeleteMakeExpenseInfo(string MakeNo, string MakeNoLn, string Seq)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                Cursor = Cursors.Default;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE ");
                strSql.AppendLine("   FROM MAKE_EXPENSE ");
                strSql.AppendLine("  WHERE MAKENO = " + MakeNo + " ");
                strSql.AppendLine("    AND MAKENO_LN = " + MakeNoLn + " ");
                strSql.AppendLine("    AND LN_ESEQ = " + Seq + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                Dispose();

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }
        
        private void RepoTxtNumber_Leave(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            string sQuantity = GridViewRetr.GetFocusedRowCellValue("EQTY")?.ToString();
            string sUnitPrc = GridViewRetr.GetFocusedRowCellValue("EDAN")?.ToString();

            double dQuantity = string.IsNullOrEmpty(sQuantity) ? 0 : Convert.ToDouble(sQuantity);
            double dUnitPrc = string.IsNullOrEmpty(sUnitPrc) ? 0 : Convert.ToDouble(sUnitPrc);

            double dResult = dQuantity * dUnitPrc;

            Cursor = Cursors.Default;

            GridViewRetr.SetFocusedRowCellValue("EAMT", dResult);
        }

        private void RepoTxtUnitPrc_Leave(object sender, EventArgs e)
        {
            RepoTxtNumber_Leave(null, null);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Dispose();
        }

        private void RepoTxtUnitPrc_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue("EDAN", txt.EditValue);
        }

        private void RepoTxtNumber_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue("EQTY", txt.EditValue);
        }

        private void RepoTxtTotAmt_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue("EQTY", 1);
            GridViewRetr.SetFocusedRowCellValue("EDAN", txt.EditValue);
            GridViewRetr.SetFocusedRowCellValue("EAMT", txt.EditValue);
        }
    }
}