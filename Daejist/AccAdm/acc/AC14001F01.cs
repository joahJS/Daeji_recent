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

namespace AccAdm
{
    public partial class AC14001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC14001F01()
        {
            InitializeComponent();
        }

        private void AC14001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = "2020-07-01";
            DateEditTo.EditValue = "2020-07-01";
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
            
            GridRetr.DataSource = GetInfo(dicParams);
            if(GridViewRetr.RowCount == 0)
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
            //if(dicParams["CONFIRM"].Equals("Y"))
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


            strSql.AppendLine(" SELECT CASE WHEN ROWNUM = 1 THEN X2.J_DATE ELSE '' END J_DATE1 ");
            strSql.AppendLine("      , CASE WHEN ROWNUM = 1 THEN X2.ATGUB ELSE '' END ATGUB1 ");
            strSql.AppendLine("      , X2.J_ID ");
            strSql.AppendLine("      , X2.CLOSE_GB ");
            strSql.AppendLine("      , X2.J_DATE ");
            strSql.AppendLine("      , X2.K_JGUBUN ");
            strSql.AppendLine("      , X2.ATGUB ");
            strSql.AppendLine("      , X2.K_JUNPYOID ");
            strSql.AppendLine("      , X2.ACCOD ");
            strSql.AppendLine("      , X2.ACNAM ");
            strSql.AppendLine("      , X2.CVCOD ");
            strSql.AppendLine("      , X2.CVNAM ");
            strSql.AppendLine("      , X2.SLIP_RMK ");
            strSql.AppendLine("      , X2.ACAMT ");
            strSql.AppendLine("      , X2.ADAMT ");
            strSql.AppendLine("      , X2.YN ");
            strSql.AppendLine("      , X2.J_CHECK ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT ROW_NUMBER () OVER(PARTITION BY A.K_JunpyoID, A.J_DATE ORDER BY A.K_JUNPYOID, A.J_DATE, A.J_ID) AS ROWNUM ");
            strSql.AppendLine("               , A.J_ID ");
            strSql.AppendLine("               , A.CLOSE_GB ");
            strSql.AppendLine("               , A.J_DATE ");
            strSql.AppendLine("               , A.K_JGUBUN ");
            strSql.AppendLine("               , A.ATGUB ");
            strSql.AppendLine("               , A.K_JUNPYOID ");
            strSql.AppendLine("               , A.ACCOD ");
            strSql.AppendLine("               , A.ACNAM ");
            strSql.AppendLine("               , A.CVCOD ");
            strSql.AppendLine("               , B.DEALER_NM AS CVNAM ");
            strSql.AppendLine("               , CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1 , ')')END)) AS SLIP_RMK ");
            strSql.AppendLine("               , A.K_CHA AS ACAMT ");
            strSql.AppendLine("               , A.K_DAE AS ADAMT ");
            strSql.AppendLine("               , CASE WHEN B1.TDATE IS NOT NULL THEN 'Y' ELSE 'N' END AS YN ");
            strSql.AppendLine("               , CASE WHEN A.J_CHECK = 'Y' THEN 'Y' ELSE 'N' END AS J_CHECK ");
            strSql.AppendLine("            FROM JUNPYO_MIG_STEP2 A ");
            strSql.AppendLine("            LEFT JOIN ACC_DEALER_CD B  ");
            strSql.AppendLine("              ON A.CVCOD = B.DEALER_CD ");
            strSql.AppendLine("            LEFT OUTER JOIN COM_BASE_CD G  ");
            strSql.AppendLine("              ON A.ATGUB = G.COM_CD  ");
            strSql.AppendLine("             AND G.CD_GB = 'AC01001_03'  ");
            strSql.AppendLine("            LEFT OUTER JOIN ACMSTF C  ");
            strSql.AppendLine("              ON A.ACCOD = C.ACCOD  ");
            strSql.AppendLine("            LEFT JOIN ACTRAN B1  ");
            strSql.AppendLine("              ON REPLACE(A.J_DATE, '-', '') = B1.TDATE ");
            strSql.AppendLine("             AND A.ACCOD = B1.ACCOD ");
            strSql.AppendLine("             AND A.CVCOD = B1.CVCOD ");
            strSql.AppendLine("             AND ISNULL(A.K_CHA, 0) = ISNULL(B1.ACAMT, 0) ");
            strSql.AppendLine("             AND ISNULL(A.K_DAE, 0) = ISNULL(B1.ADAMT, 0) ");
            strSql.AppendLine("           WHERE A.J_DATE BETWEEN '" + dicParams["YMD_F"] + "' AND '" + dicParams["YMD_T"] + "' ");
            strSql.AppendLine("             AND A.CLOSE_GB = '" + dicParams["CLOSE_GB"] + "' ");
            if (dicParams["CONFIRM"].Equals("Y"))
                strSql.AppendLine("         AND A.J_CHECK = 'Y' ");
            if (dicParams["CONFIRM"].Equals("N"))
                strSql.AppendLine("         AND (A.J_CHECK IS NULL OR A.J_CHECK <> 'Y') ");
            strSql.AppendLine("             AND (('" + dicParams["FIND_WORD"] + "' = '' AND 1 = 1 ) ");
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
            strSql.AppendLine("                 ('" + dicParams["FIND_IDX"] + "' = '6' AND B.DEALER_NM LIKE '%" + dicParams["FIND_WORD"] + "%') )");
            strSql.AppendLine("        ) X2");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                AC14001F02 frm = new AC14001F02();
                frm.Owner = this;
                frm.PARENT_FORM = this;
                frm.AddModifyGb = "ADD";
                frm.K_JUNPYOID = GridViewRetr.GetFocusedRowCellValue(GridColK_JUNPYOID)?.ToString();
                //frm.ShowDialog(this);
                //frm.Show();
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr_Click(null, null);
                }
            }
        }

        private void BtnMig_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("YMD_F", DateEditFrom.EditValue?.ToString());
            dicParams.Add("YMD_T", DateEditTo.EditValue?.ToString());
            dicParams.Add("CLOSE_GB", RdgbMigYn.EditValue?.ToString());
            dicParams.Add("FIND_IDX", CboFindSbj.SelectedIndex.ToString());
            dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());
            dicParams.Add("CONFIRM", RdgbConfirm.EditValue?.ToString());

            AC14001F03 frm = new AC14001F03();
            frm.DATE_F = DateEditFrom.EditValue?.ToString();
            frm.DATE_T = DateEditTo.EditValue?.ToString();
            frm.FIND_IDX = CboFindSbj.SelectedIndex;
            frm.FIND_WORD = TxtFindWord.EditValue?.ToString().Trim();
            frm.MIG_YN = RdgbMigYn.EditValue?.ToString();
            frm.CONFIRM = RdgbConfirm.EditValue?.ToString();
            frm.Show();
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
            if(e.KeyCode == Keys.F3)
            {
                BtnMig.PerformClick();
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if(e.KeyCode == Keys.Escape)
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
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_06'");
            }
            else if (sGb.Equals("5"))
            {
                strSql.AppendLine(" SELECT EMP_ID ");
                strSql.AppendLine("      , EMP_NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER (ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
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
            strSql.AppendLine(" WHERE A.J_DATE BETWEEN CONVERT(DATETIME, '" + sYmdFrom + "') AND CONVERT(DATETIME,'" + sYmdTo + "')                   ");
            strSql.AppendLine(" GROUP BY CONCAT(A.K_JUKYO, (CASE WHEN A.K_JUKYO1 = '' THEN '' ELSE CONCAT(' (', A.K_JUKYO1, ')')END))   ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void RdgbConfirm_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}