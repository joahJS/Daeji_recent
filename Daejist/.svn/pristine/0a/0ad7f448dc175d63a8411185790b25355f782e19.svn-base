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

/*
 * 작성일자 : 2020-02월 초
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-07
 * 수정자 : 고혜성
 * 수정내용 : 거래처초성검색 추가 (쿼리 참조)
 */
namespace AccAdm
{
    public partial class AC02001F03 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F03()
        {
            InitializeComponent();
        }

        public string DealerCd { get; set; }
        public DataRow SelectedRow { get; set; }

        public AC02001F02 P_AC02001F02;
        public AC07001F02 P_AC07001F02;
        public AC09001F01 P_AC09001F01;
        public AC14001F02 P_AC14001F02;
        public AC16001F01 P_AC16001F01;
        public AC18001F02 P_AC18001F02;

        public string ACNT_YN;
        private void AC02001F03_Load(object sender, EventArgs e)
        {
            //CboFindWord_SelectedIndexChanged(null, null);
            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);
            if (!string.IsNullOrEmpty(DealerCd))
            {
                TxtDealer.EditValue = DealerCd; 
                BtnRetr.PerformClick();
            }

            if (P_AC16001F01 != null)
            {
                BtnRetr.PerformClick();
            }
        }

        private void AC02001F03_Shown(object sender, EventArgs e)
        {
            if(GridViewRetr.RowCount > 0)
            {
                GridViewRetr.Focus();
            }
        }

        private DataTable GetLookUpEditData (string sGb, string sNullYn)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '전체' AS NM ");
                strSql.AppendLine("      , 0 AS SEQ");
                strSql.AppendLine("  UNION ALL ");
            }
            else if (sNullYn.Equals("N"))
            {
                strSql.AppendLine(" SELECT '전체' AS CD ");
                strSql.AppendLine("      , '' AS NM ");
                strSql.AppendLine("      , 0 AS SEQ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT DEALER_CD AS CD     ");
                strSql.AppendLine(" 	 , '' AS NM     ");
                strSql.AppendLine("      , DEALER_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEALER_CD");
                strSql.AppendLine("  GROUP BY DEALER_NM         ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT DEALER_NM AS CD     ");
                strSql.AppendLine(" 	 , '' AS NM     ");
                strSql.AppendLine("      , DEALER_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEALER_CD");
                strSql.AppendLine("  GROUP BY DEALER_NM         ");
            }

            strSql.AppendLine("  ORDER BY SEQ ");
            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void CboFindWord_SelectedIndexChanged(object sender, EventArgs e)
        {
            //int selIndex = CboFindWord.SelectedIndex;

            //if (selIndex == 0)
            //{
            //    DataTable dt = GetLookUpEditData("1", "N");
            //    ComGrid.SetLookUpEdit(LkupDealer, dt, "CD", "CD", "");
            //    LkupDealer.Properties.PopulateColumns();
            //    LkupDealer.Properties.Columns[1].Visible = false;
            //}
            //else if (selIndex == 1)
            //{
            //    DataTable dt = GetLookUpEditData("2", "N");
            //    ComGrid.SetLookUpEdit(LkupDealer, dt, "CD", "CD", "");
            //    LkupDealer.Properties.PopulateColumns();
            //    LkupDealer.Properties.Columns[1].Visible = false;
            //}
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sFindWord = TxtDealer.EditValue?.ToString();
            DataTable dt = GetDealerInfo(sFindWord);
            GridRetr.DataSource = dt;
            if(dt.Rows.Count > 0)
            {
                GridViewRetr.Focus();
                GridViewRetr.FocusedColumn = GridColDealerCd;
                GridViewRetr.FocusedRowHandle = 0;
                //GridViewRetr.SelectRow(0);
            }
            else
            {
                TxtDealer.SelectAll();
                TxtDealer.Focus();
            }

        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            SubmitRowDate();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {

        }
        
        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                SubmitRowDate();
            }
        }
        
        private DataTable GetDealerInfo(string sFindWord)
        {
            StringBuilder strSql = new StringBuilder();

            /*
             * 수정일자 : 2021-02-07 (현업요청)
             * 수정자 : 고혜성
             * 수정내용 : 거래처초성검색 추가
             */
            strSql.Clear();
            strSql.AppendLine(" SELECT A.DEALER_CD  ");
            strSql.AppendLine(" 	 , A.DEALER_NM  ");
            strSql.AppendLine(" 	 , A.IDT_NO  ");
            strSql.AppendLine(" 	 , A.DEALER_GB  ");
            strSql.AppendLine("      , CASE WHEN A.ADDR IS NULL OR REPLACE(A.ADDR, ' ', '') = '' THEN DTL_ADDR ELSE ADDR END AS ADDR ");
            strSql.AppendLine("      , A.REP_NM ");
            strSql.AppendLine("      , A.BIZ_NM ");
            strSql.AppendLine("      , A.TYPE_NM ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A  ");
            strSql.AppendLine("  WHERE 1 = 1 ");
            strSql.AppendLine("    AND A.EOB_YN = 'N' ");
            //은행제외
            //strSql.AppendLine("    AND A.DEALER_GB NOT IN('계좌', '금융')    ");
            //strSql.AppendLine("    AND A.DEALER_CD NOT BETWEEN 8999 AND 10000");
            if (!string.IsNullOrEmpty(sFindWord))
            {
                if(CboFindWord.SelectedIndex == 0)
                    strSql.AppendLine("    AND A.DEALER_CD = " + sFindWord);
                else if (CboFindWord.SelectedIndex == 1)
                    strSql.AppendLine("    AND (A.DEALER_NM LIKE '%" + sFindWord + "%' OR A.INITIAL_NM LIKE '%" + sFindWord + "%' )");
            }
            if(P_AC16001F01 != null)
            {
                if (!string.IsNullOrEmpty(ACNT_YN)) //계정조회일 경우 DELAER_GB 계좌로 띄움
                {
                    if (ACNT_YN.Equals("Y"))
                    {
                        strSql.AppendLine("    AND A.BANKYN = 'Y'");
                    }
                }

            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void SubmitRowDate()
        {
            DataRow row = GridViewRetr.GetFocusedDataRow();

            if (P_AC02001F02 != null)
                P_AC02001F02.DrDealerInfo = row;
            else if (P_AC07001F02 != null)
                P_AC07001F02.DrDealerInfo = row;
            else if (P_AC09001F01 != null)
                P_AC09001F01.DrDealerInfo = row;
            else if (P_AC14001F02 != null)
                P_AC14001F02.DrDealerInfo = row;
            else if (P_AC16001F01 != null)
                P_AC16001F01.DrDealerInfo = row;
            else if(P_AC18001F02 != null)
                P_AC18001F02.DrDealerInfo = row;

            DialogResult = DialogResult.OK;
        }

        private void AC02001F03_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                BtnClose_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSelect_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        #region[GridView Row's Stripe Pattern]
        
        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        #endregion[GridView Row's Stripe Pattern]

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSelect_Click(null, null);
        }

        private void TxtDealer_KeyDown(object sender, KeyEventArgs e)
        {

        }

       
    }
}