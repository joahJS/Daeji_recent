using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Net;
using System.Net.Sockets;
using System.Data.SqlClient;
/*
 * 작성일자: 모름
 * 작성자: 모름
 * 
 * ---------------- 이력----------------
 * 수정일자:2021-08-27
 * 수정자: 정은영
 * 참조번호: #0001
 * 수정내용: IP 변수 추가
 */
namespace ComLib
{
    public static class ClsFunc
    {
        #region [ 사업자번호 유효성 체크 ]

        public static bool ValidChkBsnNo(string pData)
        {
            int sum = 0;
            int sidliy = 0;
            int sidchk = 0;

            string[] getlist = new string[10];
            string[] chkvalue = new string[] { "1", "3", "7", "1", "3", "7", "1", "3", "5" };

            pData = pData.Replace("_", "").Trim();

            if (pData.Length != 10)
                return false;

            for (int i = 0; i < 10; i++)
            {
                //getlist[i] = Convert.ToString(pData.Substring(i, i + 1));
                getlist[i] = Convert.ToString(pData.Substring(i, 1));
            }
            for (int i = 0; i < 9; i++)
            {
                sum += Convert.ToInt32(getlist[i]) * Convert.ToInt32(chkvalue[i]);
            }
            sum = sum + (Convert.ToInt32(getlist[8]) * 5 / 10);
            sidliy = sum % 10;
            sidchk = 0;

            if (sidliy != 0)
            {
                sidchk = 10 - sidliy;
            }
            else
            {
                sidchk = 0;
            }
            if (sidchk != Convert.ToInt32(getlist[9])) { return false; }

            return true;
        }

        #endregion [ 사업자번호 유효성 체크 ]

        #region [ DateEdit From-To 날짜 비교 ]

        public static bool ValidChkFromToRetrYmd(DevExpress.XtraEditors.DateEdit editFr, DevExpress.XtraEditors.DateEdit editTo)
        {
            bool bRtn = true;
            string sYmdFrom = editFr.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = editTo.EditValue.ToString().Replace("-", "").Substring(0, 8);

            if (string.IsNullOrEmpty(sYmdFrom) || string.IsNullOrEmpty(sYmdTo))
                return false;

            double dFrYmd = Convert.ToDouble(sYmdFrom);
            double dToYmd = Convert.ToDouble(sYmdTo);

            if (dToYmd < dFrYmd) bRtn = false;

            return bRtn;
        }

        #endregion [ DateEdit From-To 날짜 비교 ]

        #region [관리항목 유형 관리항목 코드 조회 ]

        public static DataTable GetMgmtCd(string sMgmtGb)
        {
            DataTable dt = new DataTable();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine(" SELECT CD  ");
            strSql.AppendLine("      , NM ");
            strSql.AppendLine("   FROM V_MGMT_CD ");
            strSql.AppendLine("  WHERE MGMT_GB = '" + sMgmtGb + "' ");
            strSql.AppendLine("  ORDER BY CD   ");

            //strSql.AppendLine(" SELECT BANK_ACNT_NO AS CD  ");
            //strSql.AppendLine("      , BANK_ACNT_NO AS NM ");
            //strSql.AppendLine("   FROM ACC_ACNT_CD ");
            //strSql.AppendLine("  WHERE 'BANK_ACNT_NO' = '" + sMgmtGb + "' ");
            //strSql.AppendLine("  UNION ALL  ");
            //strSql.AppendLine(" SELECT PRNOTE_NO AS CD  ");
            //strSql.AppendLine("      , PRNOTE_NO AS NM ");
            //strSql.AppendLine("   FROM ACC_PRNOTE_MGT ");
            //strSql.AppendLine("  WHERE PRNOTE_GB = '1' ");
            //strSql.AppendLine("    AND 'PRNOTE_NO_RECEIVE' = '" + sMgmtGb + "' ");
            //strSql.AppendLine("  UNION ALL  ");
            //strSql.AppendLine(" SELECT PRNOTE_NO AS CD  ");
            //strSql.AppendLine("      , PRNOTE_NO AS NM ");
            //strSql.AppendLine("   FROM ACC_PRNOTE_MGT ");
            //strSql.AppendLine("  WHERE PRNOTE_GB = '2' ");
            //strSql.AppendLine("    AND 'PRNOTE_NO_PAY' = '" + sMgmtGb + "'   ");
            //strSql.AppendLine("  ORDER BY CD   ");

            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion [관리항목 유형 관리항목 코드 조회 ]

        #region [관리항목 코드 조회 ]

        public static DataTable GetMgmtByMgmtCd(string sMgmtGb, string sMgmtCd)
        {
            DataTable dt = new DataTable();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine(" SELECT CD  ");
            strSql.AppendLine("      , NM ");
            strSql.AppendLine("   FROM V_MGMT_CD ");
            strSql.AppendLine("  WHERE MGMT_GB = '" + sMgmtGb + "' ");
            strSql.AppendLine("    AND CD = '" + sMgmtCd + "' ");
            strSql.AppendLine("  ORDER BY CD   ");

            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion [관리항목 코드 조회 ]

        #region [계정 기본정보 조회 ]

        public static DataTable GetAccDefaultInfo(string sAccCd)
        {
            DataTable dt = new DataTable();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT ACC_CD ");
            strSql.AppendLine("      , ACC_NM ");
            strSql.AppendLine("      , UP_ACC_CD ");
            strSql.AppendLine("      , (SELECT X.ACC_NM FROM ACC_ACC_CD X WHERE X.ACC_CD = A.UP_ACC_CD) AS UP_ACC_NM ");
            strSql.AppendLine("      , DBCR_GB  ");
            strSql.AppendLine("      , PYBC_YN ");
            strSql.AppendLine("      , EVDN_ESSN_YN ");
            strSql.AppendLine("      , DEALER_YN ");
            strSql.AppendLine("      , DEBT_MGMT_GB1 AS DEBT_MGMT_GB ");
            strSql.AppendLine("      , (SELECT X.COM_NM FROM COM_BASE_CD X WHERE X.CD_GB = 'MGMT_GB' AND X.COM_CD = A.DEBT_MGMT_GB1) AS DEBT_MGMT_GB_NM ");
            strSql.AppendLine("      , CRDT_MGMT_GB1 AS CRDT_MGMT_GB ");
            strSql.AppendLine("      , (SELECT X.COM_NM FROM COM_BASE_CD X WHERE X.CD_GB = 'MGMT_GB' AND X.COM_CD = A.CRDT_MGMT_GB1) AS CRDT_MGMT_GB_NM ");
            strSql.AppendLine("      , CPTL_CD ");
            strSql.AppendLine("      , RPLC_ACC_CD ");
            strSql.AppendLine("   FROM ACC_ACC_CD A ");
            strSql.AppendLine("  WHERE ACC_CD = '" + sAccCd + "' ");

            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion [ 계정 기본정보 조회 ]

        #region [ 거래처정보 조회 사업자등록번호 ]

        public static DataTable GetDealerInfoByIdtNo(string sIdtNo)
        {
            DataTable dt = new DataTable();

            sIdtNo = sIdtNo.Replace("-", "");

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT DEALER_CD ");
            strSql.AppendLine("      , IDT_NO ");
            strSql.AppendLine("      , STRT_YMD ");
            strSql.AppendLine("      , END_YMD ");
            strSql.AppendLine("      , DEALER_GB ");
            strSql.AppendLine("      , DEALER_NM ");
            strSql.AppendLine("      , INITIAL_NM ");
            strSql.AppendLine("      , REP_NM ");
            strSql.AppendLine("      , BIZ_NM ");
            strSql.AppendLine("      , TYPE_NM ");
            strSql.AppendLine("      , ZIP ");
            strSql.AppendLine("      , ZIP_SEQ ");
            strSql.AppendLine("      , ADDR ");
            strSql.AppendLine("      , DTL_ADDR ");
            strSql.AppendLine("      , BANK_CD ");
            strSql.AppendLine("      , BANK_ACNT_NO ");
            strSql.AppendLine("      , ACNT_HOLDER ");
            strSql.AppendLine("      , EMAIL ");
            strSql.AppendLine("      , HOMEPG ");
            strSql.AppendLine("      , FAX ");
            strSql.AppendLine("      , WEB_FAX_YN ");
            strSql.AppendLine("      , SCM_PSWD ");
            strSql.AppendLine("      , CHRG_ID ");
            strSql.AppendLine("      , CHRG_NM ");
            strSql.AppendLine("      , CHRG_TEL_NO ");
            strSql.AppendLine("      , CHRG_HP_NO ");
            strSql.AppendLine("      , CHRG_EMAIL ");
            strSql.AppendLine("      , BFR_DEALER_CD ");
            strSql.AppendLine("      , BFR_DEALER_NM ");
            strSql.AppendLine("      , BILL_KIND ");
            strSql.AppendLine("      , BILL_ISSUE_GB ");
            strSql.AppendLine("      , PAY_GB ");
            strSql.AppendLine("      , PAY_TERM ");
            strSql.AppendLine("      , SPPL_LIMIT_DCNT ");
            strSql.AppendLine("      , EOB_YN ");
            strSql.AppendLine("      , NOTE ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A ");
            strSql.AppendLine("  WHERE IDT_NO = '" + sIdtNo + "' ");

            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion [ 거래처정보 조회 사업자등록번호 ]

        #region [ 소득자 정보 조회 식별번호 ]

        public static DataTable GetIncmRByIdtNo(string sName)
        {
            DataTable dt = new DataTable();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT INCMR_CD ");
            strSql.AppendLine("      , INCMR_GB ");
            strSql.AppendLine("      , IFNULL(CONCAT(LEFT(INCMR_IDT_NO, 4), '**-*******'),  CONCAT(LEFT(PP_NO, 4), '******')) AS INCMR_IDT_NO ");
            strSql.AppendLine("      , INCMR_NM ");
            strSql.AppendLine("      , RESID_NTN ");
            strSql.AppendLine("      , CO_NM ");
            strSql.AppendLine("      , BZRG_NO ");
            strSql.AppendLine("      , BZPLC_LOC ");
            strSql.AppendLine("      , ZIP ");
            strSql.AppendLine("      , ZIP_SEQ ");
            strSql.AppendLine("      , ADDR ");
            strSql.AppendLine("      , DTL_ADDR ");
            strSql.AppendLine("      , TEL_NO ");
            strSql.AppendLine("      , EMAIL ");
            strSql.AppendLine("      , BIZ_NM ");
            strSql.AppendLine("      , BANK_CD ");
            strSql.AppendLine("      , BANK_ACNT_NO ");
            strSql.AppendLine("      , ACNT_HOLDER ");
            strSql.AppendLine("      , USE_YN ");
            strSql.AppendLine("      , NOTE ");
            strSql.AppendLine("   FROM ACC_INCM_CD ");
            strSql.AppendLine("  WHERE INCMR_NM LIKE '%" + sName + "% ");

            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion [ 소득자 정보 조회 식별번호 ]

        #region [ 사원 정보 조회  ]

        public static DataTable GetEmpInfoByEmpId(string sEmpId)
        {
            DataTable dt = new DataTable();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT EMP_ID ");
            strSql.AppendLine("      , EMP_NM ");
            strSql.AppendLine("      , ENG_NM ");
            strSql.AppendLine("      , CHN_NM ");
            strSql.AppendLine("      , ENTRANCE_YMD ");
            strSql.AppendLine("      , RETIRE_YMD ");
            strSql.AppendLine("      , DEPT_CD ");
            strSql.AppendLine("      , REAL_DUTY_DEPT ");
            strSql.AppendLine("      , MANAGER_ID ");
            strSql.AppendLine("      , JOBKIND_CD ");
            strSql.AppendLine("      , JOBPOSITION_CD ");
            strSql.AppendLine("      , GRADE_CD ");
            strSql.AppendLine("      , JOBDUTY_CD ");
            strSql.AppendLine("      , PAYSTEP_CD ");
            strSql.AppendLine("      , GENDER_CD ");
            strSql.AppendLine("      , IDT_NO ");
            strSql.AppendLine("      , CNTR_GB ");
            strSql.AppendLine("      , RETIRE_RESN ");
            strSql.AppendLine("      , RETIRE_RESN_CD ");
            strSql.AppendLine("      , ENTRANCE_GB ");
            strSql.AppendLine("      , EMPL_GB ");
            strSql.AppendLine("      , CHG_BFR_ID1 ");
            strSql.AppendLine("      , CHG_BFR_ID2 ");
            strSql.AppendLine("      , FOREIGNER_YN ");
            strSql.AppendLine("   FROM HR_EMP_BASIS ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion [ 사원 정보 조회  ]

        #region [GridView Layout Setting]

        public static void SetGridViewLayout(string sId, string sProject, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            //string sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sId + @"\" + sClass + ".xaml";\
            string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            if (!File.Exists(sFile))
            {
                //sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sClass + ".xml";
                sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
                if (!File.Exists(sFile)) return;
            }
            view.RestoreLayoutFromXml(sFile);
        }

        #endregion [GridView Layout Setting]

        #region [GridView Layout Save]

        public static void SaveGridViewLayout(string sId, string sProject, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            //string path = @"C:\STLNT\" + sProject + @"\xaml\" + sId;
            string path = Application.StartupPath + @"\xaml\" + sId;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //string sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            string sFile = Application.StartupPath + @"\xaml\" + sId + @"\" + sClass + ".xaml";
            view.SaveLayoutToXml(sFile);
        }

        #endregion [GridView Layout Save]
        
        #region [GridView Layout Reset]

        public static bool ResetGridViewLayout(string sProject, string sClass, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            string sFile = @"C:\STLNT\" + sProject + @"\xaml\" + sClass + ".xml";

            if (!File.Exists(sFile)) return false;

            view.RestoreLayoutFromXml(sFile);
            return true;
        }

        #endregion [GridView Layout Reset]

        #region [Program Authority Info]
                                                        
        public static DataRow GetAutorityInfo(DataTable dtUserAuthorInfo, string sProgramName)
        {
            string sPgmNm = sProgramName;

            DataRow rowUserAutInfo = null;
            for (int i = 0; i < dtUserAuthorInfo.Rows.Count; i++)
            {
                if (dtUserAuthorInfo.Rows[i]["PGMID"].ToString().Equals(sPgmNm))
                {
                    rowUserAutInfo = dtUserAuthorInfo.Rows[i];
                }
            }
            return rowUserAutInfo;
        }

        #endregion [Program Authority Info]

        #region [Program CRUD Authority Check]

        public static bool CheckCRUDAuthority(string sCRUD, DataRow rowUserAuthorInfo)
        {
            if (rowUserAuthorInfo == null)
            {
                MessageBox.Show("사용자 권한이 없습니다.");
                return false;
            }

            string sUse = rowUserAuthorInfo["USE_Y"]?.ToString();
            string sAdd = rowUserAuthorInfo["ADD_Y"]?.ToString();
            string sUpt = rowUserAuthorInfo["UPD_Y"]?.ToString();
            string sDel = rowUserAuthorInfo["DEL_Y"]?.ToString();
            string sPrt = rowUserAuthorInfo["PRT_Y"]?.ToString();
            string sXls = rowUserAuthorInfo["XLS_Y"]?.ToString();

            bool bChkYn = false;
            if (sCRUD.Equals("READ"))
            {
                if (sUse.Equals("Y"))
                {
                    bChkYn = true;
                }
                else if (sUse.Equals("N"))
                {
                    bChkYn = false;
                }
            }
            else if (sCRUD.Equals("ADD"))
            {
                if (sAdd.Equals("Y"))
                {
                    bChkYn = true;
                }
                else if(sAdd.Equals("N"))
                {
                    bChkYn = false;
                }
            }
            else if (sCRUD.Equals("UPDATE"))
            {
                if (sUpt.Equals("Y"))
                {
                    bChkYn = true;
                }
                else if (sUpt.Equals("N"))
                {
                    bChkYn = false;
                }
            }
            else if (sCRUD.Equals("DELETE"))
            {
                if (sDel.Equals("Y"))
                {
                    bChkYn = true;
                }
                else if (sDel.Equals("N"))
                {
                    bChkYn = false;
                }
            }
            else if (sCRUD.Equals("PRINT"))
            {
                if (sPrt.Equals("Y"))
                {
                    bChkYn = true;
                }
                else if (sPrt.Equals("N"))
                {
                    bChkYn = false;
                }
            }
            else if (sCRUD.Equals("EXCEL"))
            {
                if (sXls.Equals("Y"))
                {
                    bChkYn = true;
                }
                else if (sXls.Equals("N"))
                {
                    bChkYn = false;
                }
            }
            else
            {
                bChkYn = false;
            }

            return bChkYn;
        }

        #endregion [Program CRUD Authority Check]

        #region[ENT_ID & MFY_ID Info]
        public static string GetUserIdOfLoginInfo(string row)
        {
            string sUsrCd = row.ToString();

            if (string.IsNullOrEmpty(sUsrCd))
            {
                return "X";
            }
            else
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.USRID ");
                strSql.AppendLine("   FROM ZUSRLST A ");
                strSql.AppendLine("  WHERE USRCD = " + sUsrCd + " ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                string sUsrId = dt?.Rows[0]["USRID"].ToString();

                if (string.IsNullOrEmpty(sUsrId))
                {
                    return "X";
                }
                else
                {
                    return sUsrId;
                }
            }
        }

        #endregion[ENT_ID & MFY_ID Info]

        #region[Insert Log]
        public static void LogInsert(string sDateTime, string sUsrCd, string sLogSeq, string sEditKind, string sPgmId, string sRemark, SqlCommand cmd)
        {
            string sIp = GetLocalIP();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "DP_IST_LOG";

            cmd.Parameters.Clear();
            //IP 가져오기 #0001
            cmd.Parameters.Add("@_OCCUR_DATE", SqlDbType.VarChar).Value = sDateTime;
            cmd.Parameters.Add("@_USRCD", SqlDbType.VarChar).Value = sUsrCd;
            cmd.Parameters.Add("@_LOG_SEQ", SqlDbType.VarChar).Value = sLogSeq;
            cmd.Parameters.Add("@_EDIT_KIND", SqlDbType.VarChar).Value = sEditKind;
            cmd.Parameters.Add("@_PGM_ID", SqlDbType.VarChar).Value = sPgmId;
            cmd.Parameters.Add("@_ACS_IP", SqlDbType.VarChar).Value = sIp;
            cmd.Parameters.Add("@_EDIT_RMK", SqlDbType.VarChar).Value = sRemark;

            cmd.ExecuteNonQuery();
            
        }
        #endregion[Insert Log]
        
        #region[Get IP]
        public static string GetLocalIP()
        {
            string localIP = "";
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    localIP = ip.ToString();
                    break;
                }
            }
            return localIP;
        }
        #endregion[Get IP]
    }
}
