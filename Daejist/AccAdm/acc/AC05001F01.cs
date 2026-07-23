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
using System.Diagnostics;
using ComLib;
using DevExpress.XtraGrid.Views.Grid;
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
*/
namespace AccAdm
{
    public partial class AC05001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC05001F01()
        {
            InitializeComponent();
        }
        
        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC05001F01_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            DateEditFrom.EditValue = today.AddDays(1 - today.Day);

            DataTable dtAcCod = GetLookUpData("2", "Y", "Y"); //전표구분
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupAcntCd, dtAcCod, GridBal, BGridColAcntCd, "CD", "NM", "");
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { advBandedGridView1 };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }


            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 6);
            GridBal.DataSource = GetSlipsDailyBalance(sYmdFrom);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ComnEtcFunc.ExportExcelFile("일계표_", GridBal);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AC04001F01_KeyDown(object sender, KeyEventArgs e)
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
                BtnExcel_Click(null, null);
            }
        }

        #region[Execute By Query]

        private DataTable GetSlipsDailyBalance(string sYmd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT ACCOD ");
            strSql.AppendLine("      , ACDSP ");
            strSql.AppendLine("      , 0 AS CTOT");
            strSql.AppendLine("      , SUM(IFNULL(ACAMT,0)) AS CAAMT ");
            strSql.AppendLine("      , SUM(IFNULL(HCAMT,0)) AS CHAMT ");
            strSql.AppendLine("      , 0 AS DTOT");
            strSql.AppendLine("      , SUM(IFNULL(ADAMT,0)) AS DAAMT ");
            strSql.AppendLine("      , SUM(IFNULL(HDAMT,0)) AS DHAMT ");
            strSql.AppendLine("  FROM( ");
            strSql.AppendLine("        SELECT A1.ACCOD AS ACCOD ");
            strSql.AppendLine("             , A1.ACDSP ");
            strSql.AppendLine("             , SUM(CASE WHEN A2.ASMCD <> '0101' THEN ACAMT END) AS ACAMT ");
            strSql.AppendLine("             , SUM(CASE WHEN A2.ASMCD = '0101' THEN ACAMT END) AS HCAMT ");
            strSql.AppendLine("             , SUM(CASE WHEN A2.ASMCD <> '0101' THEN ADAMT END) AS ADAMT ");
            strSql.AppendLine("             , SUM(CASE WHEN A2.ASMCD = '0101' THEN ADAMT END) AS HDAMT ");
            strSql.AppendLine("          FROM ACTOPF A1 ");
            strSql.AppendLine("          LEFT OUTER JOIN ACTRAN A2 ON ");
            strSql.AppendLine("                          A2.ACCOD BETWEEN A1.AFROM AND A1.ATO ");
            strSql.AppendLine("                      AND A2.TDATE BETWEEN '" + sYmd + "01' AND '" + sYmd + "31' ");
            strSql.AppendLine("                      AND A2.ACCOD <> '0101' ");
            strSql.AppendLine("         GROUP BY A1.ACCOD, A1.SEQNO, A1.ACDSP ");
            strSql.AppendLine("         UNION ALL ");
            strSql.AppendLine("        SELECT A1.ACCOD ");
            strSql.AppendLine("             , B1.ACDSP ");
            strSql.AppendLine("             , (CASE WHEN A1.ASMCD <> '0101' THEN ACAMT END) AS ACAMT ");
            strSql.AppendLine("             , (CASE WHEN A1.ASMCD = '0101' THEN ACAMT END) AS HCAMT ");
            strSql.AppendLine("             , (CASE WHEN A1.ASMCD <> '0101' THEN ADAMT END) AS ADAMT ");
            strSql.AppendLine("             , (CASE WHEN A1.ASMCD = '0101' THEN ADAMT END) AS HDAMT ");
            strSql.AppendLine("          FROM ACTRAN A1 ");
            strSql.AppendLine("          LEFT OUTER JOIN ACMSTF B1 ");
            strSql.AppendLine("            ON A1.ACCOD = B1.ACCOD ");
            strSql.AppendLine("         WHERE A1.TDATE BETWEEN '" + sYmd + "01' AND '" + sYmd + "31' AND A1.ACCOD <> '0101' ");
            strSql.AppendLine("      ) AS P1 ");
            strSql.AppendLine("  WHERE IFNULL(P1.ACAMT, 0) <> 0 OR IFNULL(P1.HCAMT, 0) <> 0 ");
            strSql.AppendLine("     OR IFNULL(P1.ADAMT, 0) <> 0 OR IFNULL(P1.HDAMT, 0) <> 0 ");
            strSql.AppendLine("  GROUP BY P1.ACCOD, P1.ACDSP ");
            strSql.AppendLine("  ORDER BY P1.ACCOD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sChaAlt = dt.Rows[i]["CAAMT"]?.ToString();
                string sChaCash = dt.Rows[i]["CHAMT"]?.ToString();

                double dChaAlt = string.IsNullOrEmpty(sChaAlt) ? 0 : Convert.ToDouble(sChaAlt);
                double dChaCash = string.IsNullOrEmpty(sChaCash) ? 0 : Convert.ToDouble(sChaCash);

                string sDaeAlt = dt.Rows[i]["DAAMT"]?.ToString();
                string sDaeCash = dt.Rows[i]["DHAMT"]?.ToString();

                double dDaeAlt = string.IsNullOrEmpty(sDaeAlt) ? 0 : Convert.ToDouble(sDaeAlt);
                double dDaeCash = string.IsNullOrEmpty(sDaeCash) ? 0 : Convert.ToDouble(sDaeCash);

                dt.Rows[i]["CTOT"] = dChaAlt + dChaCash;
                dt.Rows[i]["DTOT"] = dDaeAlt + dDaeCash;

            }

            return dt;
        }



        #endregion[Execute By Query]

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

            if (sGb.Equals("1")) //전표구분
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC02001_01 '");
            }
            if (sGb.Equals("2")) //계정코드
            {
                strSql.AppendLine(" SELECT ACCOD AS CD  ");
                strSql.AppendLine("      , ACNAM AS NM  ");
                strSql.AppendLine("      , ACCOD AS SEQ ");
                strSql.AppendLine("   FROM ACMSTF A  ");
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

        #region[GridView's Design]

        private void AdvBGridViewBal_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AdvBGridViewBal_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        #endregion[GridView's Design]

        private void AC05001F01_TextChanged(object sender, EventArgs e)
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

        private void DateEditFrom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}