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
    public partial class AC14001F03 : DevExpress.XtraEditors.XtraForm
    {
        public AC14001F03()
        {
            InitializeComponent();
        }

        public string DATE_F, DATE_T;
        public int FIND_IDX;
        public string FIND_WORD;
        public string MIG_YN;
        public string CONFIRM;

        private void AC14001F03_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateEditFrom.EditValue = DATE_F;
            DateEditTo.EditValue = DATE_T;
            CboFindSbj.SelectedIndex = FIND_IDX;
            TxtFindWord.EditValue = FIND_WORD;
            RdgbMigYn.EditValue = MIG_YN;
            RdgbConfirm.EditValue = CONFIRM;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("YMD_F", DateEditFrom.EditValue?.ToString());
            dicParams.Add("YMD_T", DateEditTo.EditValue?.ToString());
            dicParams.Add("CLOSE_GB", RdgbMigYn.EditValue?.ToString());
            dicParams.Add("FIND_IDX", CboFindSbj.SelectedIndex.ToString());
            dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());
            dicParams.Add("CONFIRM", RdgbConfirm.EditValue?.ToString());

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);


            GridRetr.DataSource = GetInfo(dicParams);
            GridSlip.DataSource = GetSlipInfo(sYmdFrom, sYmdTo, ReturningByComboBoxValues(CboFindSbj.EditValue?.ToString(), TxtFindWord.EditValue?.ToString().Replace("-", "")), "0");
            if (GridViewRetr.RowCount == 0)
            {
                DateEditFrom.Focus();
            }
            else
            {
                GridViewRetr.Focus();
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColJ_CHECK, "N");
            }
        }

        private DataTable GetInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");

            #region mariaDB
            //strSql.AppendLine(" SELECT CASE WHEN ROWNUM = 1 THEN X2.J_DATE ELSE '' END J_DATE1 ");
            //strSql.AppendLine("      , CASE WHEN ROWNUM = 1 THEN X2.ATGUB ELSE '' END ATGUB1 ");
            //strSql.AppendLine("      , X2.J_ID ");
            //strSql.AppendLine("      , X2.CLOSE_GB ");
            //strSql.AppendLine("      , X2.J_DATE ");
            //strSql.AppendLine("      , X2.K_JGUBUN ");
            //strSql.AppendLine("      , X2.ATGUB ");
            //strSql.AppendLine("      , X2.K_JUNPYOID ");
            //strSql.AppendLine("      , X2.ACCOD ");
            //strSql.AppendLine("      , X2.ACNAM ");
            //strSql.AppendLine("      , X2.CVCOD ");
            //strSql.AppendLine("      , X2.CVNAM ");
            //strSql.AppendLine("      , X2.SLIP_RMK ");
            //strSql.AppendLine("      , X2.ACAMT ");
            //strSql.AppendLine("      , X2.ADAMT ");
            //strSql.AppendLine("      , X2.YN ");
            //strSql.AppendLine("      , X2.J_CHECK ");
            //strSql.AppendLine("   FROM ( ");
            //strSql.AppendLine("         SELECT (CASE WHEN @VJOB = X1.K_JUNPYOID AND @TDATE = X1.J_DATE THEN @ROWNUM:=@ROWNUM+1 ELSE @ROWNUM:=1 END) ROWNUM ");
            //strSql.AppendLine("              , (@VJOB := X1.K_JUNPYOID) VJOB ");
            //strSql.AppendLine("              , (@TDATE := X1.J_DATE) VTDATE ");
            //strSql.AppendLine("              , X1.* ");
            //strSql.AppendLine("           FROM ");
            //strSql.AppendLine("              ( ");
            //strSql.AppendLine("                  SELECT A.J_ID ");
            //strSql.AppendLine("                       , A.CLOSE_GB ");
            //strSql.AppendLine("                       , A.J_DATE ");
            //strSql.AppendLine("                       , A.K_JGUBUN ");
            //strSql.AppendLine("                       , A.ATGUB ");
            //strSql.AppendLine("                       , A.K_JUNPYOID ");
            //strSql.AppendLine("                       , A.ACCOD ");
            //strSql.AppendLine("                       , A.ACNAM ");
            //strSql.AppendLine("                       , A.CVCOD ");
            //strSql.AppendLine("                       , B.DEALER_NM AS CVNAM ");
            //strSql.AppendLine("                       , CONCAT(A.K_JUKYO, IF(A.K_JUKYO1 = '', '', CONCAT(' (', A.K_JUKYO1 , ')'))) AS SLIP_RMK ");
            //strSql.AppendLine("                       , A.K_CHA AS ACAMT ");
            //strSql.AppendLine("                       , A.K_DAE AS ADAMT ");
            //strSql.AppendLine("                       , CASE WHEN B1.TDATE IS NOT NULL THEN 'Y' ELSE 'N' END AS YN ");
            //strSql.AppendLine("                       , CASE WHEN A.J_CHECK = 'Y' THEN 'Y' ELSE 'N' END AS J_CHECK ");
            //strSql.AppendLine("                    FROM JUNPYO_MIG_STEP2 A ");
            //strSql.AppendLine("                    LEFT JOIN ACC_DEALER_CD B  ");
            //strSql.AppendLine("                      ON A.CVCOD = B.DEALER_CD ");
            //strSql.AppendLine("                    LEFT OUTER JOIN COM_BASE_CD G  ");
            //strSql.AppendLine("                      ON A.ATGUB = G.COM_CD  ");
            //strSql.AppendLine("                     AND G.CD_GB = 'AC01001_03'  ");
            //strSql.AppendLine("                    LEFT OUTER JOIN ACMSTF C  ");
            //strSql.AppendLine("                      ON A.ACCOD = C.ACCOD  ");
            //strSql.AppendLine("                    LEFT JOIN ACTRAN B1  ");
            //strSql.AppendLine("                      ON REPLACE(A.J_DATE, '-', '') = B1.TDATE ");
            //strSql.AppendLine("                     AND A.ACCOD = B1.ACCOD ");
            //strSql.AppendLine("                     AND A.CVCOD = B1.CVCOD ");
            //strSql.AppendLine("                     AND IFNULL(A.K_CHA, 0) = IFNULL(B1.ACAMT, 0) ");
            //strSql.AppendLine("                     AND IFNULL(A.K_DAE, 0) = IFNULL(B1.ADAMT, 0) ");
            //strSql.AppendLine("                   WHERE A.J_DATE BETWEEN '" + dicParams["YMD_F"] + "' AND '" + dicParams["YMD_T"] + "' ");
            //strSql.AppendLine("                     AND A.CLOSE_GB = '" + dicParams["CLOSE_GB"] + "' ");
            //if (dicParams["CONFIRM"].Equals("Y"))
            //    strSql.AppendLine("                     AND A.J_CHECK = 'Y' ");
            //if (dicParams["CONFIRM"].Equals("N"))
            //    strSql.AppendLine("                     AND (A.J_CHECK IS NULL OR A.J_CHECK <> 'Y') ");
            //strSql.AppendLine("                     AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '0' AND A.ACCOD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '1' AND A.ACNAM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '2' AND G.COM_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '3' AND C.ASMCD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '4' AND A.ACCOD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '5' AND CONCAT(A.K_JUKYO, IF(A.K_JUKYO1 = '', '', CONCAT(' (', A.K_JUKYO1 , ')'))) LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            //strSql.AppendLine("                          OR ");
            //strSql.AppendLine("                          ('" + dicParams["FIND_IDX"] + "' = '6' AND B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ) ");
            //strSql.AppendLine("                   ORDER BY A.K_JUNPYOID, A.J_DATE, A.J_ID ");
            //strSql.AppendLine("              ) X1, (SELECT @ROWNUM := 0) R ");
            //strSql.AppendLine("         ) X2 ");
            #endregion

            strSql.AppendLine("SELECT CASE WHEN ROWNUM = 1 THEN X2.J_DATE ELSE '' END J_DATE1                                                                   ");
            strSql.AppendLine("     , CASE WHEN ROWNUM = 1 THEN X2.ATGUB ELSE '' END ATGUB1                                                                     ");
            strSql.AppendLine("     , X2.J_ID                                                                                                                   ");
            strSql.AppendLine("     , X2.CLOSE_GB                                                                                                               ");
            strSql.AppendLine("     , X2.J_DATE                                                                                                                 ");
            strSql.AppendLine("     , X2.K_JGUBUN                                                                                                               ");
            strSql.AppendLine("     , X2.ATGUB                                                                                                                  ");
            strSql.AppendLine("     , X2.K_JUNPYOID                                                                                                             ");
            strSql.AppendLine("     , X2.ACCOD                                                                                                                  ");
            strSql.AppendLine("     , X2.ACNAM                                                                                                                  ");
            strSql.AppendLine("     , X2.CVCOD                                                                                                                  ");
            strSql.AppendLine("     , X2.CVNAM                                                                                                                  ");
            strSql.AppendLine("     , X2.SLIP_RMK                                                                                                               ");
            strSql.AppendLine("     , X2.ACAMT                                                                                                                  ");
            strSql.AppendLine("     , X2.ADAMT                                                                                                                  ");
            strSql.AppendLine("     , X2.YN                                                                                                                     ");
            strSql.AppendLine("     , X2.J_CHECK                                                                                                                ");
            strSql.AppendLine("  FROM(                                                                                                                          ");
            strSql.AppendLine("         SELECT ROW_NUMBER() OVER(PARTITION BY A.K_JunpyoID, A.J_DATE ORDER BY A.K_JUNPYOID, A.J_DATE, A.J_ID) AS ROWNUM         ");
            strSql.AppendLine("              , A.J_ID                                                                                                           ");
            strSql.AppendLine("              , A.CLOSE_GB                                                                                                       ");
            strSql.AppendLine("              , A.J_DATE                                                                                                         ");
            strSql.AppendLine("              , A.K_JGUBUN                                                                                                       ");
            strSql.AppendLine("              , A.ATGUB                                                                                                          ");
            strSql.AppendLine("              , A.K_JUNPYOID                                                                                                     ");
            strSql.AppendLine("              , A.ACCOD                                                                                                          ");
            strSql.AppendLine("              , A.ACNAM                                                                                                          ");
            strSql.AppendLine("              , A.CVCOD                                                                                                          ");
            strSql.AppendLine("              , B.DEALER_NM AS CVNAM                                                                                             ");
            strSql.AppendLine("              , CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1, ')')END)) AS SLIP_RMK         ");
            strSql.AppendLine("              , A.K_CHA AS ACAMT                                                                                                 ");
            strSql.AppendLine("              , A.K_DAE AS ADAMT                                                                                                 ");
            strSql.AppendLine("              , CASE WHEN B1.TDATE IS NOT NULL THEN 'Y' ELSE 'N' END AS YN                                                       ");
            strSql.AppendLine("              , CASE WHEN A.J_CHECK = 'Y' THEN 'Y' ELSE 'N' END AS J_CHECK                                                       ");
            strSql.AppendLine("           FROM JUNPYO_MIG_STEP2 A ");
            strSql.AppendLine("           LEFT JOIN ACC_DEALER_CD B  ");
            strSql.AppendLine("             ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendLine("           LEFT OUTER JOIN COM_BASE_CD G  ");
            strSql.AppendLine("             ON A.ATGUB = G.COM_CD  ");
            strSql.AppendLine("            AND G.CD_GB = 'AC01001_03'  ");
            strSql.AppendLine("           LEFT OUTER JOIN ACMSTF C  ");
            strSql.AppendLine("             ON A.ACCOD = C.ACCOD  ");
            strSql.AppendLine("           LEFT JOIN ACTRAN B1  ");
            strSql.AppendLine("             ON REPLACE(A.J_DATE, '-', '') = B1.TDATE ");
            strSql.AppendLine("            AND A.ACCOD = B1.ACCOD ");
            strSql.AppendLine("            AND A.CVCOD = B1.CVCOD ");
            strSql.AppendLine("            AND ISNULL(A.K_CHA, 0) = ISNULL(B1.ACAMT, 0) ");
            strSql.AppendLine("            AND ISNULL(A.K_DAE, 0) = ISNULL(B1.ADAMT, 0) ");
            strSql.AppendLine("          WHERE A.J_DATE BETWEEN '" + dicParams["YMD_F"] + "' AND '" + dicParams["YMD_T"] + "' ");
            strSql.AppendLine("            AND A.CLOSE_GB = '" + dicParams["CLOSE_GB"] + "' ");
            if (dicParams["CONFIRM"].Equals("Y"))
                strSql.AppendLine("        AND A.J_CHECK = 'Y' ");
            if (dicParams["CONFIRM"].Equals("N"))
                strSql.AppendLine("        AND (A.J_CHECK IS NULL OR A.J_CHECK <> 'Y') ");
            strSql.AppendLine("            AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '0' AND A.ACCOD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '1' AND A.ACNAM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '2' AND G.COM_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '3' AND C.ASMCD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '4' AND A.ACCOD LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '5' AND CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1 , ')')END)) LIKE '%" + dicParams["FIND_WORD"] + "%') ");
            strSql.AppendLine("                 OR ");
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '6' AND B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') ) ");
            strSql.AppendLine("        ) X2");


            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetSlipInfo(string sYmdFrom, string sYmdTo, string AddingQuery, string sAprvGb)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT CASE WHEN ROWNUM = 1 THEN X2.TDATE ELSE '' END TDATE1 ");
            //strSql.AppendLine("      , CASE WHEN ROWNUM = 1 THEN X2.ATGUB ELSE '' END ATGUB1 ");
            //strSql.AppendLine("      , CASE WHEN ROWNUM = 1 THEN X2.SEQNO ELSE '' END SEQNO1 ");
            //strSql.AppendLine("      , X2.TDATE   ");
            //strSql.AppendLine("      , X2.ATGUB  ");
            //strSql.AppendLine("      , X2.SEQNO   ");
            //strSql.AppendLine("      , X2.LINNO  ");
            //strSql.AppendLine("      , X2.ACCOD  ");
            //strSql.AppendLine("      , X2.ACNAM  ");
            //strSql.AppendLine("      , X2.CVCOD   ");
            //strSql.AppendLine("      , X2.CVNAM  ");
            //strSql.AppendLine("      , X2.ACTCD   ");
            //strSql.AppendLine("      , X2.ACTNM   ");
            //strSql.AppendLine("      , X2.ATEXT  ");
            //strSql.AppendLine("      , X2.ACAMT  ");
            //strSql.AppendLine("      , X2.ADAMT  ");
            //strSql.AppendLine("      , X2.ADPCD  ");
            //strSql.AppendLine("      , X2.ADPNM   ");
            //strSql.AppendLine("      , X2.APVYN ");
            //strSql.AppendLine("      , X2.AAUTO  ");
            //strSql.AppendLine("      , X2.ADATE  ");
            //strSql.AppendLine("      , X2.AUSER  ");
            //strSql.AppendLine("      , X2.BILNO  ");
            //strSql.AppendLine("      , X2.RK  ");
            //strSql.AppendLine("      , X2.CUSER  ");
            //strSql.AppendLine("      , X2.CDATE  ");
            //strSql.AppendLine("      , X2.MUSER  ");
            //strSql.AppendLine("      , X2.MDATE  ");
            //strSql.AppendLine("      , X2.REF1  ");
            //strSql.AppendLine("      , X2.REF2  ");
            //strSql.AppendLine("      , X2.REF3  ");
            //strSql.AppendLine("   FROM(  ");
            //strSql.AppendLine("        SELECT (CASE WHEN @VJOB = X1.SEQNO AND @TDATE = X1.TDATE THEN @ROWNUM:=@ROWNUM+1 ELSE @ROWNUM:=1 END) ROWNUM ");
            //strSql.AppendLine("             , (@VJOB := X1.SEQNO) VJOB ");
            //strSql.AppendLine("             , (@TDATE := X1.TDATE) VTDATE ");
            //strSql.AppendLine("             , X1.* ");
            //strSql.AppendLine("          FROM ");
            //strSql.AppendLine("             ( ");
            //strSql.AppendLine("               SELECT DATE_FORMAT(A.TDATE,'%Y-%m-%d') AS TDATE   ");
            //strSql.AppendLine("         	        , G.COM_NM AS ATGUB  ");
            //strSql.AppendLine("         	        , A.SEQNO   ");
            //strSql.AppendLine("         	        , A.LINNO  ");
            //strSql.AppendLine("         	        , A.ACCOD  ");
            //strSql.AppendLine("         	        , B.ACNAM AS ACNAM  ");
            //strSql.AppendLine("         	        , CAST(A.CVCOD AS CHAR) AS CVCOD   ");
            //strSql.AppendLine("         	        , A.CVNAM  ");
            //strSql.AppendLine("         	        , A.ACTCD   ");
            //strSql.AppendLine("         	        , D.ACNAM AS ACTNM   ");
            //strSql.AppendLine("         	        , A.ATEXT  ");
            //strSql.AppendLine("         	        , A.ACAMT  ");
            //strSql.AppendLine("         	        , A.ADAMT  ");
            //strSql.AppendLine("         	        , A.ADPCD  ");
            //strSql.AppendLine("         	        , E.DEPT_NM AS ADPNM   ");
            //strSql.AppendLine("         	        , A.REF1  ");
            //strSql.AppendLine("         	        , A.REF2  ");
            //strSql.AppendLine("         	        , A.REF3  ");
            //strSql.AppendLine("                     , A.APVYN ");
            //strSql.AppendLine("         	        , A.AAUTO  ");
            //strSql.AppendLine("         	        , A.ADATE  ");
            //strSql.AppendLine("         	        , A.AUSER  ");
            //strSql.AppendLine("                     , F.BILNO  ");
            //strSql.AppendLine("         	        , A.RK  ");
            //strSql.AppendLine("         	        , CAST(A.CUSER AS CHAR) AS CUSER  ");
            //strSql.AppendLine("         	        , A.CDATE  ");
            //strSql.AppendLine("         	        , CAST(A.MUSER AS CHAR) AS MUSER  ");
            //strSql.AppendLine("         	        , A.MDATE  ");
            //strSql.AppendLine("                 FROM ACTRAN A ");
            //strSql.AppendLine("                 LEFT OUTER JOIN ACMSTF B  ");
            //strSql.AppendLine("                   ON A.ACCOD = B.ACCOD  ");
            //strSql.AppendLine("                 LEFT OUTER JOIN ACC_DEALER_CD C   ");
            //strSql.AppendLine("                   ON A.CVCOD = C.DEALER_CD  ");
            //strSql.AppendLine("                 LEFT OUTER JOIN ACMSTF D  ");
            //strSql.AppendLine("                   ON A.ACTCD = D.ACCOD  ");
            //strSql.AppendLine("                 LEFT OUTER JOIN ACC_DEPT_CD E  ");
            //strSql.AppendLine("                   ON A.ADPCD = E.DEPT_CD  ");
            //strSql.AppendLine("                 LEFT OUTER JOIN ACBILL F  ");
            //strSql.AppendLine("                   ON A.TDATE = F.TDATE  ");
            //strSql.AppendLine("                  AND A.ATGUB = F.ATGUB   ");
            //strSql.AppendLine("                  AND A.SEQNO = F.SEQNO   ");
            //strSql.AppendLine("                  AND A.LINNO = F.LINNO   "); 
            //strSql.AppendLine("                 LEFT OUTER JOIN COM_BASE_CD G  ");
            //strSql.AppendLine("                   ON A.ATGUB = G.COM_CD  ");
            //strSql.AppendLine("                  AND G.CD_GB = 'AC02001_01'  ");
            //strSql.AppendLine("                WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("                  AND (('" + sAprvGb + "' = '0' AND 1 = 1)");
            //strSql.AppendLine("                       OR ");
            //strSql.AppendLine("                       ('" + sAprvGb + "' = '1' AND A.APVYN = 'Y' ) ");
            //strSql.AppendLine("                       OR ");
            //strSql.AppendLine("                       ('" + sAprvGb + "' = '2' AND A.APVYN <> 'Y' )) ");
            //strSql.AppendLine(AddingQuery);
            //strSql.AppendLine("                ORDER BY A.TDATE, A.SEQNO, A.LINNO ");
            //strSql.AppendLine("             ) X1, (SELECT @ROWNUM := 0) R ");
            //strSql.AppendLine("        ) X2 ");
            #endregion

            strSql.AppendLine("SELECT CASE WHEN ROWNUM = 1 THEN X2.TDATE ELSE NULL END TDATE1 ");
            strSql.AppendLine("     , CASE WHEN ROWNUM = 1 THEN X2.ATGUB ELSE NULL END ATGUB1 ");
            strSql.AppendLine("     , CASE WHEN ROWNUM = 1 THEN X2.SEQNO ELSE NULL END SEQNO1 ");
            strSql.AppendLine("     , X2.TDATE                                              ");
            strSql.AppendLine("     , X2.ATGUB                                              ");
            strSql.AppendLine("     , X2.SEQNO                                              ");
            strSql.AppendLine("     , X2.LINNO                                              ");
            strSql.AppendLine("     , X2.ACCOD                                              ");
            strSql.AppendLine("     , X2.ACNAM                                              ");
            strSql.AppendLine("     , X2.CVCOD                                              ");
            strSql.AppendLine("     , X2.CVNAM                                              ");
            strSql.AppendLine("     , X2.ACTCD                                              ");
            strSql.AppendLine("     , X2.ACTNM                                              ");
            strSql.AppendLine("     , X2.ATEXT                                              ");
            strSql.AppendLine("     , X2.ACAMT                                              ");
            strSql.AppendLine("     , X2.ADAMT                                              ");
            strSql.AppendLine("     , X2.ADPCD                                              ");
            strSql.AppendLine("     , X2.ADPNM                                              ");
            strSql.AppendLine("     , X2.APVYN                                              ");
            strSql.AppendLine("     , X2.AAUTO                                              ");
            strSql.AppendLine("     , X2.ADATE                                              ");
            strSql.AppendLine("     , X2.AUSER                                              ");
            strSql.AppendLine("     , X2.BILNO                                              ");
            strSql.AppendLine("     , X2.RK                                                 ");
            strSql.AppendLine("     , X2.CUSER                                              ");
            strSql.AppendLine("     , X2.CDATE                                              ");
            strSql.AppendLine("     , X2.MUSER                                              ");
            strSql.AppendLine("     , X2.MDATE                                              ");
            strSql.AppendLine("     , X2.REF1                                               ");
            strSql.AppendLine("     , X2.REF2                                               ");
            strSql.AppendLine("     , X2.REF3                                               ");
            strSql.AppendLine("  FROM(                                                      ");
            strSql.AppendLine("          SELECT ROW_NUMBER() OVER(PARTITION BY A.TDATE, A.ATGUB, A.SEQNO ORDER BY A.TDATE, A.SEQNO, A.LINNO) AS ROWNUM");
            strSql.AppendLine("               , CONVERT(DATE, A.TDATE) AS TDATE  ");
            strSql.AppendLine("               , G.COM_NM AS ATGUB                ");
            strSql.AppendLine("               , A.SEQNO                          ");
            strSql.AppendLine("               , A.LINNO                          ");
            strSql.AppendLine("               , A.ACCOD                          ");
            strSql.AppendLine("               , B.ACNAM AS ACNAM                 ");
            strSql.AppendLine("               , A.CVCOD AS CVCOD                 ");
            strSql.AppendLine("               , A.CVNAM                          ");
            strSql.AppendLine("               , A.ACTCD                          ");
            strSql.AppendLine("               , D.ACNAM AS ACTNM                 ");
            strSql.AppendLine("               , A.ATEXT                          ");
            strSql.AppendLine("               , A.ACAMT                          ");
            strSql.AppendLine("               , A.ADAMT                          ");
            strSql.AppendLine("               , A.ADPCD                          ");
            strSql.AppendLine("               , E.DEPT_NM AS ADPNM               ");
            strSql.AppendLine("               , A.REF1                           ");
            strSql.AppendLine("               , A.REF2                           ");
            strSql.AppendLine("               , A.REF3                           ");
            strSql.AppendLine("               , A.APVYN                          ");
            strSql.AppendLine("               , A.AAUTO                          ");
            strSql.AppendLine("               , A.ADATE                          ");
            strSql.AppendLine("               , A.AUSER                          ");
            strSql.AppendLine("               , F.BILNO                          ");
            strSql.AppendLine("               , A.RK                             ");
            strSql.AppendLine("               , A.CUSER AS CUSER                 ");
            strSql.AppendLine("               , A.CDATE                          ");
            strSql.AppendLine("               , A.MUSER AS MUSER                 ");
            strSql.AppendLine("               , A.MDATE                          ");
            strSql.AppendLine("            FROM ACTRAN A                         ");
            strSql.AppendLine("            LEFT OUTER JOIN ACMSTF B              ");
            strSql.AppendLine("              ON A.ACCOD = B.ACCOD                ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C       ");
            strSql.AppendLine("              ON A.CVCOD = C.DEALER_CD            ");
            strSql.AppendLine("            LEFT OUTER JOIN ACMSTF D              ");
            strSql.AppendLine("              ON A.ACTCD = D.ACCOD                ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEPT_CD E         ");
            strSql.AppendLine("              ON A.ADPCD = E.DEPT_CD              ");
            strSql.AppendLine("            LEFT OUTER JOIN ACBILL F              ");
            strSql.AppendLine("              ON A.TDATE = F.TDATE                ");
            strSql.AppendLine("             AND A.ATGUB = F.ATGUB                ");
            strSql.AppendLine("             AND A.SEQNO = F.SEQNO                ");
            strSql.AppendLine("             AND A.LINNO = F.LINNO                ");
            strSql.AppendLine("            LEFT OUTER JOIN COM_BASE_CD G         ");
            strSql.AppendLine("              ON A.ATGUB = G.COM_CD               ");
            strSql.AppendLine("             AND G.CD_GB = 'AC02001_01'           ");
            strSql.AppendLine("                WHERE A.TDATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            strSql.AppendLine("                  AND (('" + sAprvGb + "' = '0' AND 1 = 1)");
            strSql.AppendLine("                       OR ");
            strSql.AppendLine("                       ('" + sAprvGb + "' = '1' AND A.APVYN = 'Y' ) ");
            strSql.AppendLine("                       OR ");
            strSql.AppendLine("                       ('" + sAprvGb + "' = '2' AND A.APVYN <> 'Y' )) ");
            strSql.AppendLine(AddingQuery);
            strSql.AppendLine("       ) X2                                                ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DataTable dtSave = GetK_JunpyoID();

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                foreach (DataRow row in dtSave.Rows)
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT REPLACE(J_DATE, '-', '') AS TDATE ");
                    strSql.AppendLine("      , ATGUB ");
                    strSql.AppendLine("      , '' AS SEQNO ");
                    strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY A.J_ID) AS LINNO ");
                    strSql.AppendLine("      , A.ACCOD ");
                    strSql.AppendLine("      , A.ACNAM ");
                    strSql.AppendLine("      , A.CVCOD ");
                    strSql.AppendLine("      , B.DEALER_NM AS CVNAM ");
                    strSql.AppendLine("      , CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1 , ')')END)) AS ATEXT ");
                    strSql.AppendLine("      , ISNULL(A.K_CHA, 0) AS ACAMT ");
                    strSql.AppendLine("      , ISNULL(A.K_DAE, 0) AS ADAMT ");
                    strSql.AppendLine("      , '전표이관' AS RK ");
                    strSql.AppendLine("   FROM JUNPYO_MIG_STEP2 A  ");
                    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD B  ");
                    strSql.AppendLine("     ON A.CVCOD = B.DEALER_CD ");
                    strSql.AppendLine("  WHERE A.K_JUNPYOID = " + row["K_JUNPYOID"] + " ");

                    DataTable dtSlip = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    string sTDate = dtSlip.Rows[0]["TDATE"]?.ToString();
                    string sAtGub = dtSlip.Rows[0]["ATGUB"]?.ToString();
                    string sSeqNo = GetSlipNo(cmd, sTDate, sAtGub);
                    foreach (DataRow slip in dtSlip.Rows)
                    {
                        string sLinNo = slip["LINNO"]?.ToString();
                        string sAcCod = slip["ACCOD"]?.ToString();
                        string sAcNam = slip["ACNAM"]?.ToString();
                        string sCvCod = slip["CVCOD"]?.ToString();
                        string sCvNam = slip["CVNAM"]?.ToString();
                        string sAText = slip["ATEXT"]?.ToString();
                        string sAcAmt = slip["ACAMT"]?.ToString();
                        string sAdAmt = slip["ADAMT"]?.ToString();
                        string sRk = slip["RK"]?.ToString();
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

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" INSERT INTO ACTRAN ");
                        strSql.AppendLine("           ( TDATE ");
                        strSql.AppendLine("           , ATGUB ");
                        strSql.AppendLine("           , SEQNO ");
                        strSql.AppendLine("           , LINNO ");
                        strSql.AppendLine("           , ACCOD ");
                        strSql.AppendLine("           , ACNAM ");
                        strSql.AppendLine("           , ATEXT ");
                        strSql.AppendLine("           , CVCOD ");
                        strSql.AppendLine("           , CVNAM ");
                        strSql.AppendLine("           , ACAMT ");
                        strSql.AppendLine("           , ADAMT ");
                        strSql.AppendLine("           , RK ");
                        strSql.AppendLine("           , REF1 ");
                        strSql.AppendLine("           , REF2 ");
                        strSql.AppendLine("           , REF3 ");
                        strSql.AppendLine("           , CUSER ");
                        strSql.AppendLine("           , ADPCD ");
                        strSql.AppendLine("           , ADPNM ");
                        strSql.AppendLine("           , CDATE )");
                        strSql.AppendLine("     VALUES( @TDATE ");
                        strSql.AppendLine("           , @ATGUB ");
                        strSql.AppendLine("           , @SEQNO ");
                        strSql.AppendLine("           , @LINNO ");
                        //strSql.AppendLine("           ,  " + sLinNo + " ");
                        strSql.AppendLine("           , @ACCOD ");
                        strSql.AppendLine("           , (SELECT ACNAM     ");
                        strSql.AppendLine("                FROM ACMSTF    ");
                        strSql.AppendLine("               WHERE ACCOD = @ACCOD )");
                        strSql.AppendLine("           , '" + sAText + "' ");
                        strSql.AppendLine("           , @CVCOD ");
                        strSql.AppendLine("           , '" + sCvNam + "' ");
                        strSql.AppendLine("           , " + sAcAmt + " ");
                        strSql.AppendLine("           , " + sAdAmt + " ");
                        strSql.AppendLine("           , '" + sRk + "' ");
                        strSql.AppendLine("           , '" + sRef1 + "' ");
                        strSql.AppendLine("           , '" + sRef2 + "' ");
                        strSql.AppendLine("           , '" + sRef3 + "' ");
                        strSql.AppendLine("           , '" + sEmpId + "' ");
                        strSql.AppendLine("           , '" + sDeptCd + "' ");
                        strSql.AppendLine("           , ( SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = '" + sDeptCd + "' ) ");
                        strSql.AppendLine("           , CONVERT(VARCHAR(19),GETDATE(),20) ) ");

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

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE JUNPYO_MIG_STEP2 ");
                        strSql.AppendLine("    SET CLOSE_GB = 'Y' ");
                        strSql.AppendLine("  WHERE K_JUNPYOID = @K_JUNPYOID ");

                        cmd.Parameters.Clear();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.AddWithValue("@K_JUNPYOID", row["K_JUNPYOID"]);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                Cursor = Cursors.Default;

                BtnRetr.PerformClick();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);

                Cursor = Cursors.Default;
            }
        }

        private DataTable GetK_JunpyoID()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT X1.K_JUNPYOID ");
            strSql.AppendLine("   FROM (");
            int[] arrSelectedIdx = GridViewRetr.GetSelectedRows();
            int iCnt = 0;
            foreach (int idx in arrSelectedIdx)
            {
                string sK_JUNPYOID = GridViewRetr.GetRowCellValue(idx, GridColK_JUNPYOID)?.ToString();

                if (iCnt++ == arrSelectedIdx.Length - 1)
                {
                    strSql.AppendLine(" SELECT " + sK_JUNPYOID + " AS K_JUNPYOID");
                }
                else
                {
                    strSql.AppendLine(" SELECT " + sK_JUNPYOID + " AS K_JUNPYOID");
                    strSql.AppendLine("  UNION ALL ");
                }
            }

            strSql.AppendLine("        ) X1 ");
            strSql.AppendLine("  GROUP BY X1.K_JUNPYOID ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
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

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            string sYn = GridViewRetr.GetRowCellValue(e.RowHandle, GridColJ_CHECK)?.ToString();
            if (string.IsNullOrEmpty(sYn))
                return;

            if (sYn.Equals("Y"))
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
            else
            {
                e.Appearance.BackColor = SystemColors.Info;
            }
        }

        private void AC14001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                BtnSave.PerformClick();
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                //BtnClose.PerformClick();
            }
        }

        private void RdgbMigYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void CboFindSbj_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                ComboBoxEdit cbo = (ComboBoxEdit)sender;
                SettingTextEditAutoComplete(cbo.EditValue?.ToString());
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
            }
        }

        private void SettingTextEditAutoComplete(string sGb)
        {
            AutoCompleteStringCollection collection = new AutoCompleteStringCollection();

            DataTable dt = new DataTable();

            if (sGb.Equals("계정코드") || sGb.Equals("관계계정코드"))
            {
                dt = GetLookUpData("2", "Y", "Y");
            }
            if (sGb.Equals("계정명") || sGb.Equals("관계계정명"))
            {
                dt = GetLookUpData("3", "Y", "Y");
            }
            if (sGb.Equals("적요"))
            {
                dt = GetRemarkInfo();
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                collection.Add(dt.Rows[i]["CD"]?.ToString());
            }

            TxtFindWord.MaskBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            TxtFindWord.MaskBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            TxtFindWord.MaskBox.AutoCompleteCustomSource = collection;
        }

        private string ReturningByComboBoxValues(string sValue, string sText)
        {
            if (string.IsNullOrEmpty(sText))
                return string.Empty;

            string[] cnWHERE = { "A.ACCOD", "A.ACNAM", "G.COM_NM", "B.ASMCD", "(SELECT ACNAM FROM ACMSTF)", "ATEXT", "C.DEALER_NM" }; string sWHERE = "";
            if (TxtFindWord.Text != "")
            {
                sWHERE = "                  AND " + cnWHERE[CboFindSbj.SelectedIndex] + " LIKE '" + "%" + TxtFindWord.Text + "%" + "' ";
            }
            return sWHERE;

            //StringBuilder strSql = new StringBuilder();

            //if (sValue.Equals("계정코드"))
            //    strSql.AppendLine("    AND A.ACCOD LIKE '" + sText + "%' ");
            //else if (sValue.Equals("계정명"))
            //    strSql.AppendLine("    AND A.ACNAM LIKE '%" + sText + "%' ");
            //else if (sValue.Equals("성격"))
            //    strSql.AppendLine("    AND G.COM_NM LIKE '%" + sText + "%");
            //else if (sValue.Equals("관계계정코드"))
            //    strSql.AppendLine("    AND A.ASMCD LIKE '%" + sText + "%' ");
            //else if (sValue.Equals("관계계정명"))
            //{
            //    strSql.AppendLine("    AND B.ASMCD IN (SELECT X1.ACCOD ");
            //    strSql.AppendLine("                      FROM ACMSTF X1 ");
            //    strSql.AppendLine("                     WHERE X1.ACNAM LIKE '%" + sText + "%' ) ");
            //}

            //return strSql.ToString();
        }

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
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_01'");
            }
            else if (sGb.Equals("2")) //계정코드
            {
                strSql.AppendLine(" SELECT A.ACCOD AS CD ");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , '' AS SEQ");
                strSql.AppendLine("   FROM ACMSTF A ");
            }
            else if (sGb.Equals("3")) //계정명
            {
                strSql.AppendLine(" SELECT A.ACNAM AS CD ");
                strSql.AppendLine("      , '' AS NM");
                strSql.AppendLine("      , '' AS SEQ");
                strSql.AppendLine("   FROM ACMSTF A ");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_06 '");
            }
            else if (sGb.Equals("5"))
            {
                strSql.AppendLine(" SELECT EMP_ID ");
                strSql.AppendLine("      , EMP_NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER (ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A ");
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

        private DataTable GetRemarkInfo()
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();

            #region mariaDB
            //strSql.AppendLine("               SELECT CONCAT(A.K_JUKYO, IF(A.K_JUKYO1 = '', '', CONCAT(' (', A.K_JUKYO1 , ')'))) AS CD ");
            //strSql.AppendLine("                 FROM JUNPYO_MIG_STEP2 A ");
            //strSql.AppendLine("                WHERE A.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("                GROUP BY A.K_JUNPYOID ");
            #endregion

            strSql.AppendLine("SELECT CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1, ')')END)) AS CD");
            strSql.AppendLine("  FROM JUNPYO_MIG_STEP2 A                                                                                ");
            strSql.AppendLine(" WHERE A.J_DATE BETWEEN CONVERT(DATETIME, '" + sYmdFrom + "') AND CONVERT(DATETIME, '" + sYmdTo + "')                  ");
            strSql.AppendLine(" GROUP BY CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1, ')')END))   ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void RdgbConfirm_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewSlip_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewRetr_RowClick_1(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                string sK_JunpyoId = GridViewRetr.GetFocusedRowCellValue(GridColK_JUNPYOID)?.ToString();
                for(int i = 0; i < GridViewRetr.RowCount; i++)
                {
                    string sCompare = GridViewRetr.GetDataRow(i)["K_JUNPYOID"]?.ToString();
                    if (sCompare.Equals(sK_JunpyoId))
                    {
                        if(GridViewRetr.IsRowSelected(i))
                            GridViewRetr.UnselectRow(i);
                        else
                            GridViewRetr.SelectRow(i);

                        //if (GridViewRetr.IsRowSelected(e.RowHandle))
                        //    GridViewRetr.UnselectRow(i);
                        //else
                        //    GridViewRetr.SelectRow(i);
                    }
                }

                //Dictionary<int, string> dicParams = new Dictionary<int, string>();

                //DataTable dt = (DataTable)GridRetr.DataSource;
                //foreach(DataRow row in dt.Rows)
                //{
                //    dicParams.Add(dt.Rows.IndexOf(row), "");
                //}
            }
        }
    }
}