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
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;
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
    public partial class IncomeScrapDailyReport : DevExpress.XtraEditors.XtraForm
    {
        public IncomeScrapDailyReport()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void IncomeScrapDailyReport_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            TxtContAmt.Properties.Mask.Culture = new System.Globalization.CultureInfo("en-US");
            TxtContTotAmt.Properties.Mask.Culture = new System.Globalization.CultureInfo("en-US");
            TxtExchangeRate.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            TxtIcmAmt.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            TxtBrokerCost.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            TxtBankCost.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            TxtSurvey.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            TxtAvgUnpr.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            TxtTotAmt.Properties.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            RepoTxtContAmt.Mask.Culture = new System.Globalization.CultureInfo("en-US");
            RepoTxtContTotAmt.Mask.Culture = new System.Globalization.CultureInfo("en-US");
            RepoTxtExchangeRate.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            RepoTxtWon.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");
            RepoTxtAvgUnpr.Mask.Culture = new System.Globalization.CultureInfo("ko-KR");

            DataTable dtDealer = GetLookUpData("1", "Y", "Y");
            DataTable dtGrade = GetLookUpData("2", "Y", "Y");
            DataTable dtPmntMethod = GetLookUpData("3", "Y", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupPmntMethod, dtPmntMethod, "CD", "NM", "");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupShipIncomeCompany, dtDealer, GridRetr, GridColTransOpCompany, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupShipIncomeCompany, dtDealer, GridRetr, GridColIcmCompCd, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGradeCd, dtGrade, GridRetr, GridColIcmGrade, "CD", "NM", "");
            //ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupPayMethod, dtPmntMethod, GridRetr, GridColIcmPmntMethod, "CD", "NM", "");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkUpDealer, dtDealer, GridDtlRetr, GridColDtlDealerCd, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkUpGrade, dtGrade, GridDtlRetr, GridColDtlIcmGrade, "CD", "NM", "");

            

            Cursor = Cursors.Default;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr, GridViewDtlRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }
            BtnRetr_Click(null, null);
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

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT CONVERT(VARCHAR,DEALER_CD) AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.J_SERIAL AS CD");
                strSql.AppendLine("      , A.GUBUN1 AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY GUBUN1) AS SEQ");
                strSql.AppendLine("   FROM JAJAE A");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE CD_GB = 'ICM_PAY_METHOD'");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE CD_GB = 'EXCHANGE_RATE' ");
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-","").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            if(string.IsNullOrEmpty(sYmdFrom) || string.IsNullOrEmpty(sYmdTo))
            {
                MessageBox.Show("검색기간을 설정하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.ICM_STL_CD ");
            strSql.AppendLine("      , A.ICM_PAY_METHOD ");
            strSql.AppendLine("      , B.COM_NM AS ICM_PAY_NM");
            strSql.AppendLine("      , A.ICM_LC_NO ");
            strSql.AppendLine("      , A.ICM_GRADE ");
            strSql.AppendLine("      , CONVERT(DATE,A.ARRV_DATE) AS ARRV_DATE ");
            strSql.AppendLine("      , A.TRANS_OPRTN_COMPANY ");
            strSql.AppendLine("      , C.DEALER_NM AS COMNM1");
            strSql.AppendLine("      , A.ICM_BL_NO ");
            strSql.AppendLine("      , A.CONT_WEIGHT_TON ");
            strSql.AppendLine("      , A.CONT_AMT ");
            strSql.AppendLine("      , A.CONT_TOTAL_AMT ");
            strSql.AppendLine("      , A.ICM_WEIGHT ");
            strSql.AppendLine("      , A.ICM_EXCHANGE_RATE ");
            strSql.AppendLine("      , A.ICM_AMT ");
            strSql.AppendLine("      , A.CUSTOM_BROKER_COST ");
            strSql.AppendLine("      , A.BANK_COST ");
            strSql.AppendLine("      , A.SURVEY ");
            strSql.AppendLine("      , A.TOTAL_AMT ");
            strSql.AppendLine("      , A.AVG_UNPR ");
            strSql.AppendLine("      , A.ICM_COMP_CD ");
            strSql.AppendLine("      , D.DEALER_NM AS COMNM2");
            strSql.AppendLine("      , A.WARE_HOUSING_WEIGHT ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
            strSql.AppendLine("   FROM INCOME_STL_MGT A ");
            strSql.AppendLine("   LEFT JOIN COM_BASE_CD B         ");
            strSql.AppendLine("     ON A.ICM_PAY_METHOD = B.COM_CD");
            strSql.AppendLine("    AND B.CD_GB = 'ICM_PAY_METHOD'");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C               ");
            strSql.AppendLine("     ON A.TRANS_OPRTN_COMPANY = CONVERT(VARCHAR, CONVERT(NUMERIC,C.DEALER_CD))");
            strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD D               ");
            strSql.AppendLine("     ON A.ICM_COMP_CD = CONVERT(VARCHAR, CONVERT(NUMERIC,D.DEALER_CD))        ");
            strSql.AppendLine("  WHERE ARRV_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND ARRV_DATE <= '" + sYmdTo + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void DateEditFrom_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void DateEditTo_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sIcmStlCd = GridViewRetr.GetFocusedRowCellValue("ICM_STL_CD")?.ToString();

            if (e.FocusedRowHandle >= 0 && !string.IsNullOrEmpty(sIcmStlCd))
            {
                Cursor = Cursors.WaitCursor;

                TxtIcmStlCd.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_STL_CD");

                LkupPmntMethod.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_PAY_METHOD");
                LkupPmntMethod.Text = GridViewRetr.GetFocusedRowCellValue("ICM_PAY_NM")?.ToString();
                TxtLcNo.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_LC_NO");

                BtnEditGradeNm.EditValue = GridViewRetr.GetFocusedRowCellDisplayText("ICM_GRADE");
                TxtGradeCd.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_GRADE");

                string sArrvDate = GridViewRetr.GetFocusedRowCellValue("ARRV_DATE")?.ToString().Substring(0,10);
                DateEditArrv.EditValue = sArrvDate;

                BtnEditShipCompNm.EditValue = GridViewRetr.GetFocusedRowCellValue("COMNM1");
                TxtShipCompCd.EditValue = GridViewRetr.GetFocusedRowCellValue("TRANS_OPRTN_COMPANY");

                TxtBLNo.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_BL_NO");
                TxtContTon.EditValue = GridViewRetr.GetFocusedRowCellValue("CONT_WEIGHT_TON");
                TxtContAmt.EditValue = GridViewRetr.GetFocusedRowCellValue("CONT_AMT");
                TxtContTotAmt.EditValue = GridViewRetr.GetFocusedRowCellValue("CONT_TOTAL_AMT");
                TxtIncomeTon.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_WEIGHT");
                TxtExchangeRate.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_EXCHANGE_RATE");
                TxtIcmAmt.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_AMT");
                TxtBrokerCost.EditValue = GridViewRetr.GetFocusedRowCellValue("CUSTOM_BROKER_COST");
                TxtBankCost.EditValue = GridViewRetr.GetFocusedRowCellValue("BANK_COST");
                TxtSurvey.EditValue = GridViewRetr.GetFocusedRowCellValue("SURVEY");
                TxtTotAmt.EditValue = GridViewRetr.GetFocusedRowCellValue("TOTAL_AMT");
                TxtAvgUnpr.EditValue = GridViewRetr.GetFocusedRowCellValue("AVG_UNPR");

                BtnEditIcmCompNm.EditValue = GridViewRetr.GetFocusedRowCellValue("COMNM2");
                TxtIcmCompCd.EditValue = GridViewRetr.GetFocusedRowCellValue("ICM_COMP_CD");

                TxtWareHWeight.EditValue = GridViewRetr.GetFocusedRowCellValue("WARE_HOUSING_WEIGHT");

                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT A.ICM_STL_CD ");
                strSql.AppendLine("      , A.ICM_STL_SEQ ");
                strSql.AppendLine("      , A.ICM_DEALER_CD ");
                strSql.AppendLine("      , A.VEHICLE_NO ");
                strSql.AppendLine("      , A.ICM_WEIGHT ");
                strSql.AppendLine("      , A.ICM_UNPR ");
                strSql.AppendLine("      , A.ICM_GRADE ");
                strSql.AppendLine("      , A.IN_WAREHOUSE_AMT ");
                strSql.AppendLine("      , A.ICM_PAYMENT ");
                strSql.AppendLine("      , A.ICM_LOSS_WEIGHT ");
                strSql.AppendLine("      , A.ICM_CARRY_COST ");
                strSql.AppendLine("      , A.ENT_DT ");
                strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
                strSql.AppendLine("      , A.MFY_DT ");
                strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
                strSql.AppendLine("   FROM INCOME_STL_DTL A ");
                strSql.AppendLine("  WHERE ICM_STL_CD = " + sIcmStlCd + "");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                GridDtlRetr.DataSource = dt;

                Cursor = Cursors.Default;
            }
        }

        private void BtnDtlAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            IncomeScrapDailyRptEditor frm = new IncomeScrapDailyRptEditor();

            DataRow[] drArrStlDtl;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                drArrStlDtl = frm.drArrResult;

                drArrStlDtl = new DataRow[frm.drArrResult.Length];
                for (int i = 0; i < frm.drArrResult.Length; i++)
                {
                    drArrStlDtl[i] = frm.drArrResult[i];
                }

                DataTable dt = (DataTable)GridDtlRetr.DataSource;
                
                for (int i = 0; i < drArrStlDtl.Length; i++)
                {
                    DataRow dr = dt.NewRow();

                    int iSeq = 0;
                    for(int j = 0; j < dt.Rows.Count; j++)
                    {
                        if(iSeq < Convert.ToInt32(dt.Rows[j]["ICM_STL_SEQ"]))
                        {
                            iSeq = Convert.ToInt32(dt.Rows[j]["ICM_STL_SEQ"]);
                        }
                    }

                    dr["ICM_STL_CD"] = TxtIcmStlCd.EditValue.ToString();
                    dr["ICM_STL_SEQ"] = iSeq + 1;
                    dr["ICM_DEALER_CD"] = drArrStlDtl[i]["DEALER_CD"];
                    dr["VEHICLE_NO"] = drArrStlDtl[i]["J_BNUM"];
                    dr["ICM_WEIGHT"] = drArrStlDtl[i]["DANJUNG"];
                    dr["ICM_UNPR"] = drArrStlDtl[i]["DANGA"];
                    dr["ICM_GRADE"] = drArrStlDtl[i]["J_SERIAL"];
                    dr["IN_WAREHOUSE_AMT"] = drArrStlDtl[i]["IKONGKEP"];
                    dr["ICM_LOSS_WEIGHT"] = drArrStlDtl[i]["HALIN"];
                    dr["ICM_CARRY_COST"] = drArrStlDtl[i]["CKONGKEP"];
                    
                    dt.Rows.Add(dr);
                }
                GridDtlRetr.DataSource = dt;

                double dInWeight = 0;
                double dAvgUnpr = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sInWeight = dt.Rows[i]["ICM_WEIGHT"].ToString();
                    dInWeight += string.IsNullOrEmpty(sInWeight) ? 0 : Convert.ToDouble(sInWeight);
                    string sAvgUnpr = dt.Rows[i]["ICM_UNPR"].ToString();
                    dAvgUnpr += string.IsNullOrEmpty(sAvgUnpr) ? 0 : Convert.ToDouble(sAvgUnpr);
                }
                TxtWareHWeight.EditValue = dInWeight;
                TxtAvgUnpr.EditValue = Math.Round((dAvgUnpr / dt.Rows.Count), 2);

                Cursor = Cursors.Default;
            }
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;

            ClearAllForm(layoutControl1);
            
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT MAX(A.ICM_STL_CD) AS MAX_VALUE ");
            strSql.AppendLine("   FROM INCOME_STL_MGT A ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
            double dMaxValue = 0;
            if (string.IsNullOrEmpty(dt.Rows[0]["MAX_VALUE"].ToString()))
            {
                dMaxValue = 1;
            }
            else
            {
                dMaxValue = Convert.ToDouble(dt.Rows[0]["MAX_VALUE"]) + 1;
            }
            TxtIcmStlCd.EditValue = dMaxValue;

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.ICM_STL_CD ");
            strSql.AppendLine("      , A.ICM_STL_SEQ ");
            strSql.AppendLine("      , A.ICM_DEALER_CD ");
            strSql.AppendLine("      , A.VEHICLE_NO ");
            strSql.AppendLine("      , A.ICM_WEIGHT ");
            strSql.AppendLine("      , A.ICM_UNPR ");
            strSql.AppendLine("      , A.ICM_GRADE ");
            strSql.AppendLine("      , A.IN_WAREHOUSE_AMT ");
            strSql.AppendLine("      , A.ICM_PAYMENT ");
            strSql.AppendLine("      , A.ICM_LOSS_WEIGHT ");
            strSql.AppendLine("      , A.ICM_CARRY_COST ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("   FROM INCOME_STL_DTL A ");
            strSql.AppendLine("  WHERE ICM_STL_CD = " + dMaxValue + " ");

            DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridDtlRetr.DataSource = dt2;

            Cursor = Cursors.Default;
        }

        public void ClearAllForm(Control Ctrl)
        {
            if (Ctrl.HasChildren)
            {
                foreach (Control ctrl in Ctrl.Controls)
                {
                    if (ctrl is DevExpress.XtraEditors.DateEdit)
                    {
                        if ((((DevExpress.XtraEditors.BaseEdit)ctrl).Name.Equals(DateEditFrom.ToString())))
                        {
                            return;
                        }
                        else if ((((DevExpress.XtraEditors.BaseEdit)ctrl).Name.Equals(DateEditTo.ToString())))
                        {
                            return;
                        }
                        else
                        {
                            (ctrl as DevExpress.XtraEditors.DateEdit).EditValue = DateTime.Today;
                        }
                    }
                    else
                    {
                        if (ctrl is DevExpress.XtraEditors.TextEdit)
                            (ctrl as DevExpress.XtraEditors.TextEdit).EditValue = string.Empty;

                        if (ctrl is DevExpress.XtraEditors.LookUpEdit)
                            (ctrl as DevExpress.XtraEditors.LookUpEdit).EditValue = string.Empty;

                        if (ctrl is DevExpress.XtraEditors.ButtonEdit)
                            (ctrl as DevExpress.XtraEditors.ButtonEdit).EditValue = string.Empty;
                    }
                    if (ctrl.HasChildren)
                        ClearAllForm(ctrl);//Recursive
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sPmntGb = LkupPmntMethod.EditValue?.ToString();
            if (string.IsNullOrEmpty(sPmntGb))
            {
                XtraMessageBox.Show("결제방식을 선택하여주세요");
                return;
            }
            string sLCNullYn = TxtLcNo.EditValue?.ToString();
            if (string.IsNullOrEmpty(sLCNullYn))
            {
                if (sPmntGb.Equals("2"))
                {
                    XtraMessageBox.Show("결제방식이 L/C인 경우 L/C 번호를 작성해주세요.");
                    return;
                }
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                DataTable dtSave;

                if (GridViewDtlRetr.RowCount == 0)
                {
                    dtSave = null;
                }
                else
                {
                    dtSave = (DataTable)GridDtlRetr.DataSource;
                }

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sId = FmMainToolBar2.drUser["USRCD"]?.ToString();

                if (GridViewDtlRetr.RowCount > 0)
                {
                    double dIcm_Stl_Cd = 0;
                    double dIcmStlSeq = 0;
                    double dIcmDealerCd = 0;
                    string sVehicleNo = string.Empty;
                    double dIcm_Weight = 0;
                    double dIcm_Unpr = 0;
                    string sIcm_Grade = string.Empty;
                    double dInWareHouseAmt = 0;
                    string sIcmPmnt = string.Empty;
                    double dIcmLossWeight = 0;
                    double dIcmCarryCost = 0;

                    for (int i = 0; i < dtSave.Rows.Count; i++)
                    {
                        dIcm_Stl_Cd = Convert.ToDouble(dtSave.Rows[i]["ICM_STL_CD"]);
                        dIcmStlSeq = Convert.ToDouble(dtSave.Rows[i]["ICM_STL_SEQ"]);
                        dIcmDealerCd = Convert.ToDouble(dtSave.Rows[i]["ICM_DEALER_CD"]);
                        sVehicleNo = dtSave.Rows[i]["VEHICLE_NO"].ToString();
                        dIcm_Weight = Convert.ToDouble(dtSave.Rows[i]["ICM_WEIGHT"]);
                        dIcm_Unpr = Convert.ToDouble(dtSave.Rows[i]["ICM_UNPR"]);
                        sIcm_Grade = dtSave.Rows[i]["ICM_GRADE"].ToString();
                        dInWareHouseAmt = Convert.ToDouble(dtSave.Rows[i]["IN_WAREHOUSE_AMT"]);
                        sIcmPmnt = dtSave.Rows[i]["ICM_PAYMENT"].ToString();
                        dIcmLossWeight = Convert.ToDouble(dtSave.Rows[i]["ICM_LOSS_WEIGHT"]);
                        dIcmCarryCost = Convert.ToDouble(dtSave.Rows[i]["ICM_CARRY_COST"]);

                        strSql.Clear();
                        strSql.AppendLine(" IF EXISTS(SELECT* FROM INCOME_STL_DTL WHERE ICM_STL_CD = "+ dIcm_Stl_Cd + " AND ICM_STL_SEQ ="+ dIcmStlSeq + " ) ");
                        strSql.AppendLine("    BEGIN                                                                          ");
                        strSql.AppendLine("           UPDATE INCOME_STL_DTL                                                   ");
                        strSql.AppendLine("              SET VEHICLE_NO = '" + sVehicleNo + "' ");
                        strSql.AppendLine("                , ICM_WEIGHT = " + dIcm_Weight + " ");
                        strSql.AppendLine("                , ICM_UNPR = " + dIcm_Unpr + " ");
                        strSql.AppendLine("                , ICM_GRADE = '" + sIcm_Grade + "' ");
                        strSql.AppendLine(" 	           , IN_WAREHOUSE_AMT = " + dInWareHouseAmt + " ");
                        strSql.AppendLine(" 	           , ICM_PAYMENT = '" + sIcmPmnt + "' ");
                        strSql.AppendLine(" 	           , ICM_LOSS_WEIGHT = " + dIcmLossWeight + " ");
                        strSql.AppendLine(" 	           , ICM_CARRY_COST = " + dIcmCarryCost + " ");
                        strSql.AppendLine(" 	           , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20) ");
                        strSql.AppendLine(" 	           , MFY_ID = '" + sId + "' ");
                        strSql.AppendLine(" WHERE ICM_STL_CD = "+ dIcm_Stl_Cd + " AND ICM_STL_SEQ = "+ dIcmStlSeq);
                        strSql.AppendLine("      END                                                                          ");
                        strSql.AppendLine(" ELSE                                                                              ");
                        strSql.AppendLine("    BEGIN                                                                          ");
                        strSql.AppendLine("INSERT INTO INCOME_STL_DTL ");
                        strSql.AppendLine("          ( ");
                        strSql.AppendLine("            ICM_STL_CD ");
                        strSql.AppendLine("          , ICM_STL_SEQ ");
                        strSql.AppendLine("          , ICM_DEALER_CD");
                        strSql.AppendLine("          , VEHICLE_NO ");
                        strSql.AppendLine("          , ICM_WEIGHT ");
                        strSql.AppendLine("          , ICM_UNPR ");
                        strSql.AppendLine("          , ICM_GRADE ");
                        strSql.AppendLine("          , IN_WAREHOUSE_AMT ");
                        strSql.AppendLine(" 	     , ICM_PAYMENT ");
                        strSql.AppendLine(" 	     , ICM_LOSS_WEIGHT ");
                        strSql.AppendLine(" 	     , ICM_CARRY_COST ");
                        strSql.AppendLine(" 	     , ENT_DT ");
                        strSql.AppendLine(" 	     , ENT_ID ");
                        strSql.AppendLine(" 	     ) ");
                        strSql.AppendLine("     VALUES ");
                        strSql.AppendLine("          ( ");
                        strSql.AppendLine("            " + dIcm_Stl_Cd + " ");
                        strSql.AppendLine(" 		 , " + dIcmStlSeq + " ");
                        strSql.AppendLine("          , " + dIcmDealerCd + " ");
                        strSql.AppendLine(" 		 , '" + sVehicleNo + "' ");
                        strSql.AppendLine(" 		 , " + dIcm_Weight + " ");
                        strSql.AppendLine(" 	     , " + dIcm_Unpr + " ");
                        strSql.AppendLine(" 		 , '" + sIcm_Grade + "' ");
                        strSql.AppendLine(" 		 , " + dInWareHouseAmt + " ");
                        strSql.AppendLine(" 		 , '" + sIcmPmnt + "'   ");
                        strSql.AppendLine(" 		 , " + dIcmLossWeight + " ");
                        strSql.AppendLine(" 		 , " + dIcmCarryCost + " ");
                        strSql.AppendLine(" 		 , CONVERT(VARCHAR(20),GETDATE(),20) ");
                        strSql.AppendLine(" 		 , '" + sId + "' ");
                        strSql.AppendLine("          ) ");
                        strSql.AppendLine("      END");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                    }
                }

                double dIcmStlCd = 0;
                string sIcmContDate = string.Empty;
                string sIcmPayMethod = string.Empty;
                string sIcmLcNo = string.Empty;
                string sIcmGrade = string.Empty;
                string sArrvDate = string.Empty;
                string sTransOprtnComp = string.Empty;
                string sIcmBlNo = string.Empty;
                double dContWeightTon = 0;
                double dContAmt = 0;
                double dContTotAmt = 0;
                double dIcmWeight = 0;
                double dIcmExchangeRate = 0;
                double dIcmAmt = 0;
                double dCustomBrokerCost = 0;
                double dBankCost = 0;
                double dSurvey = 0;
                double dTotAmt = 0;
                double dAvgUnpr = 0;
                double dIcmCompCd = 0;
                double dWareHousingWeight = 0;

                dIcmStlCd = Convert.ToDouble(TxtIcmStlCd.EditValue?.ToString());

                sIcmPayMethod = LkupPmntMethod.EditValue?.ToString();
                sIcmPayMethod = string.IsNullOrEmpty(sIcmPayMethod) ? string.Empty : sIcmPayMethod;

                sIcmLcNo = TxtLcNo.EditValue?.ToString();
                sIcmLcNo = string.IsNullOrEmpty(sIcmPayMethod) ? string.Empty : sIcmLcNo;

                sIcmGrade = TxtGradeCd.EditValue?.ToString();
                sIcmGrade = string.IsNullOrEmpty(sIcmGrade) ? string.Empty : sIcmGrade;

                sArrvDate = DateEditArrv.EditValue.ToString()?.Replace("-", "").Substring(0, 8);
                sArrvDate = string.IsNullOrEmpty(sArrvDate) ? string.Empty : sArrvDate;

                sTransOprtnComp = TxtShipCompCd.EditValue?.ToString();
                sTransOprtnComp = string.IsNullOrEmpty(sTransOprtnComp) ? string.Empty : sTransOprtnComp;

                sIcmBlNo = TxtBLNo.EditValue?.ToString();
                sIcmBlNo = string.IsNullOrEmpty(sIcmBlNo) ? string.Empty : sIcmBlNo;

                string sContWeightTon = TxtContTon.EditValue?.ToString();
                dContWeightTon = string.IsNullOrEmpty(sContWeightTon) ? 0 : Math.Round(Convert.ToDouble(sContWeightTon), 3);

                string sContAmt = TxtContAmt.EditValue?.ToString();
                dContAmt = string.IsNullOrEmpty(sContAmt) ? 0 : Math.Round(Convert.ToDouble(sContAmt), 1);
                
                string sContTotAmt = TxtContTotAmt.EditValue?.ToString();
                dContTotAmt = string.IsNullOrEmpty(sContTotAmt) ? 0 : Math.Round(Convert.ToDouble(sContTotAmt), 2);

                string sIcmWeight = TxtIncomeTon.EditValue?.ToString();
                dIcmWeight = string.IsNullOrEmpty(sIcmWeight) ? 0 : Math.Round(Convert.ToDouble(sIcmWeight), 3);

                string sIcmExchangeRate = TxtExchangeRate.EditValue?.ToString();
                dIcmExchangeRate = string.IsNullOrEmpty(sIcmExchangeRate) ? 0 : Math.Round(Convert.ToDouble(sIcmExchangeRate), 3);

                string sIcmAmt = TxtIcmAmt.EditValue?.ToString();
                dIcmAmt = string.IsNullOrEmpty(sIcmAmt) ? 0 : Convert.ToDouble(sIcmAmt);

                string sCustomBrokerCost = TxtBrokerCost.EditValue?.ToString();
                dCustomBrokerCost = string.IsNullOrEmpty(sCustomBrokerCost) ? 0 : Convert.ToDouble(sCustomBrokerCost);

                string sBankCost = TxtBankCost.EditValue?.ToString();
                dBankCost = string.IsNullOrEmpty(sBankCost) ? 0 : Convert.ToDouble(sBankCost);

                string sSurvey = TxtSurvey.EditValue?.ToString();
                dSurvey = string.IsNullOrEmpty(sSurvey) ? 0 : Convert.ToDouble(sSurvey);

                string sTotAmt = TxtTotAmt.EditValue?.ToString();
                dTotAmt = string.IsNullOrEmpty(sTotAmt) ? 0 : Convert.ToDouble(sTotAmt);

                string sAvgUnpr = TxtAvgUnpr.EditValue?.ToString();
                dAvgUnpr = string.IsNullOrEmpty(sAvgUnpr) || sAvgUnpr.Equals("NaN") ? 0 : Math.Round(Convert.ToDouble(sAvgUnpr), 2);

                string sIcmCompCd = TxtIcmCompCd.EditValue?.ToString();
                dIcmCompCd = string.IsNullOrEmpty(sIcmCompCd) ? 0 : Convert.ToDouble(sIcmCompCd);

                string sWareHWeight = TxtWareHWeight.EditValue?.ToString();
                dWareHousingWeight = string.IsNullOrEmpty(sWareHWeight) ? 0 : Convert.ToDouble(sWareHWeight);

                strSql.Clear();
                strSql.AppendLine("IF EXISTS(SELECT* FROM INCOME_STL_MGT WHERE ICM_STL_CD = '"+ dIcmStlCd + "') ");
                strSql.AppendLine("   BEGIN                                                     ");
                strSql.AppendLine("         UPDATE INCOME_STL_MGT                               ");
                strSql.AppendLine("            SET ICM_PAY_METHOD = '" + sIcmPayMethod + "' ");
                strSql.AppendLine("              , ICM_LC_NO = '" + sIcmLcNo + "' ");
                strSql.AppendLine("              , ICM_GRADE = '" + sIcmGrade + "' ");
                strSql.AppendLine("              , ARRV_DATE = '" + sArrvDate + "' ");
                strSql.AppendLine("              , TRANS_OPRTN_COMPANY = '" + sTransOprtnComp + "' ");
                strSql.AppendLine("              , ICM_BL_NO = '" + sIcmBlNo + "' ");
                strSql.AppendLine("              , CONT_WEIGHT_TON = " + dContWeightTon + " ");
                strSql.AppendLine("              , CONT_AMT = " + dContAmt + " ");
                strSql.AppendLine("              , CONT_TOTAL_AMT = " + dContTotAmt + " ");
                strSql.AppendLine("              , ICM_WEIGHT = " + dIcmWeight + " ");
                strSql.AppendLine("              , ICM_EXCHANGE_RATE = " + dIcmExchangeRate + " ");
                strSql.AppendLine("              , ICM_AMT = " + dIcmAmt + " ");
                strSql.AppendLine("              , CUSTOM_BROKER_COST = " + dCustomBrokerCost + " ");
                strSql.AppendLine("              , BANK_COST = " + dBankCost + " ");
                strSql.AppendLine("              , SURVEY = " + dSurvey + " ");
                strSql.AppendLine("              , TOTAL_AMT = " + dTotAmt + " ");
                strSql.AppendLine("              , AVG_UNPR = " + dAvgUnpr + " ");
                strSql.AppendLine("              , ICM_COMP_CD = " + dIcmCompCd + " ");
                strSql.AppendLine("              , WARE_HOUSING_WEIGHT = " + dWareHousingWeight + " ");
                strSql.AppendLine("              , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20) ");
                strSql.AppendLine("              , MFY_ID = '"+ sId + "'                                  ");
                strSql.AppendLine("          WHERE ICM_STL_CD = '"+ dIcmStlCd + "'                              ");
                strSql.AppendLine("     END                                                     ");
                strSql.AppendLine("ELSE                                                         ");
                strSql.AppendLine("   BEGIN                                                     ");
                strSql.AppendLine("         INSERT INTO INCOME_STL_MGT                          ");
                strSql.AppendLine("           (ICM_STL_CD                                       ");
                strSql.AppendLine("           , ICM_PAY_METHOD                                  ");
                strSql.AppendLine("           , ICM_LC_NO                                       ");
                strSql.AppendLine("           , ICM_GRADE                                       ");
                strSql.AppendLine("           , ARRV_DATE                                       ");
                strSql.AppendLine("           , TRANS_OPRTN_COMPANY                             ");
                strSql.AppendLine("           , ICM_BL_NO                                       ");
                strSql.AppendLine("           , CONT_WEIGHT_TON                                 ");
                strSql.AppendLine("           , CONT_AMT                                        ");
                strSql.AppendLine("           , CONT_TOTAL_AMT                                  ");
                strSql.AppendLine("           , ICM_WEIGHT                                      ");
                strSql.AppendLine("           , ICM_EXCHANGE_RATE                               ");
                strSql.AppendLine("           , ICM_AMT                                         ");
                strSql.AppendLine("           , CUSTOM_BROKER_COST                              ");
                strSql.AppendLine("           , BANK_COST                                       ");
                strSql.AppendLine("           , SURVEY                                          ");
                strSql.AppendLine("           , TOTAL_AMT                                       ");
                strSql.AppendLine("           , AVG_UNPR                                        ");
                strSql.AppendLine("           , ICM_COMP_CD                                     ");
                strSql.AppendLine("           , WARE_HOUSING_WEIGHT                             ");
                strSql.AppendLine("           , ENT_DT                                          ");
                strSql.AppendLine("           , ENT_ID                                          ");
                strSql.AppendLine("           )                                                 ");
                strSql.AppendLine("      VALUES                                                 ");
                strSql.AppendLine("           ( " + dIcmStlCd + " ");
                strSql.AppendLine("           , '" + sIcmPayMethod + "' ");
                strSql.AppendLine("           , '" + sIcmLcNo + "' ");
                strSql.AppendLine("           , '" + sIcmGrade + "' ");
                strSql.AppendLine("           , '" + sArrvDate + "' ");
                strSql.AppendLine("           , '" + sTransOprtnComp + "' ");
                strSql.AppendLine("           , '" + sIcmBlNo + "' ");
                strSql.AppendLine("           , " + dContWeightTon + " ");
                strSql.AppendLine("           , " + dContAmt + " ");
                strSql.AppendLine("           , " + dContTotAmt + " ");
                strSql.AppendLine("           , " + dIcmWeight + " ");
                strSql.AppendLine("           , " + dIcmExchangeRate + " ");
                strSql.AppendLine("           , " + dIcmAmt + " ");
                strSql.AppendLine("           , " + dCustomBrokerCost + " ");
                strSql.AppendLine("           , " + dBankCost + " ");
                strSql.AppendLine("           , " + dSurvey + " ");
                strSql.AppendLine("           , " + dTotAmt + " ");
                strSql.AppendLine("           , " + dAvgUnpr + " ");
                strSql.AppendLine("           , " + dIcmCompCd + " ");
                strSql.AppendLine("           , " + dWareHousingWeight + " ");
                strSql.AppendLine("           , CONVERT(VARCHAR(20), GETDATE(), 20) ");
                strSql.AppendLine("           , '" + sId + "' ");
                strSql.AppendLine("           )                                                 ");
                strSql.AppendLine("     END                                                     ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);

                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColIcmStlCd, dIcmStlCd.ToString());


                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sDate = DateEditFrom.EditValue?.ToString().Substring(0, 10).Replace("-","");

            ComnEtcFunc.ExportExcelFile("수입고철일보_"+ sDate, GridRetr);

            //string sIcmStlChk = TxtIcmStlCd.EditValue?.ToString();
            //double dIcmStlCd = string.IsNullOrEmpty(sIcmStlChk) ? 0 : Convert.ToDouble(sIcmStlChk);

            //StringBuilder strSql = new StringBuilder();

            //strSql.Clear();
            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT A.ICM_STL_CD ");
            //strSql.AppendLine("   FROM INCOME_STL_MGT A ");
            //strSql.AppendLine("  WHERE ICM_STL_CD = " + dIcmStlCd + " ");

            //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //if (dt.Rows.Count == 0)
            //{
            //    Cursor = Cursors.WaitCursor;

            //    XtraMessageBox.Show("현재 입력정보가 데이터베이스에 없습니다. \r\n 입력사항들을 저장버튼을 누르시고 진행해주세요.");

            //    Cursor = Cursors.Default;
            //    return;
            //}

            //Excel.Application excelApp = null;
            //Excel.Workbook wb = null;
            //Excel.Worksheet ws = null;

            //try
            //{
            //    Cursor = Cursors.WaitCursor;
                
            //    string sPmntMtd = LkupPmntMethod.Text;
            //    string sLcNo = TxtLcNo.EditValue == null ? string.Empty : TxtLcNo.EditValue.ToString();
            //    string sGrade = BtnEditGradeNm.EditValue == null ? string.Empty : BtnEditGradeNm.EditValue.ToString();
            //    string sArrvDate = DateEditArrv.EditValue?.ToString();
            //    string sTransOprtnComp = BtnEditShipCompNm.EditValue == null ? string.Empty : BtnEditShipCompNm.EditValue.ToString();
            //    string sBlNo = TxtBLNo.EditValue == null ? string.Empty : TxtBLNo.EditValue.ToString();

            //    string sContTon = TxtContTon.EditValue?.ToString();
            //    double dContTon = 0;
            //    if (!string.IsNullOrEmpty(sContTon)) dContTon = Convert.ToDouble(sContTon);

            //    string sContAmt = TxtContAmt.EditValue?.ToString();
            //    double dContAmt = 0;
            //    if (!string.IsNullOrEmpty(sContAmt)) dContAmt = Convert.ToDouble(sContAmt);

            //    string sContTotAmt = TxtContTotAmt.EditValue?.ToString();
            //    double dContTotAmt = 0;
            //    if (!string.IsNullOrEmpty(sContTotAmt)) dContTotAmt = Convert.ToDouble(sContTotAmt);

            //    string sImtTon = TxtIncomeTon.EditValue?.ToString();
            //    double dImtTon = 0;
            //    if (!string.IsNullOrEmpty(sImtTon)) dImtTon = Convert.ToDouble(sImtTon);

            //    string sExchangeRate = TxtExchangeRate.EditValue?.ToString();
            //    double dExchangeRate = 0;
            //    if (!string.IsNullOrEmpty(sExchangeRate)) dExchangeRate = Convert.ToDouble(sExchangeRate);

            //    string sIcmAmt = TxtIcmAmt.EditValue?.ToString();
            //    double dIcmAmt = 0;
            //    if (!string.IsNullOrEmpty(sIcmAmt)) dIcmAmt = Convert.ToDouble(sIcmAmt);

            //    string sBrokerCost = TxtBrokerCost.EditValue?.ToString();
            //    double dBrokerCost = 0;
            //    if (!string.IsNullOrEmpty(sBrokerCost)) dBrokerCost = Convert.ToDouble(sBrokerCost);

            //    string sBankCost = TxtBankCost.EditValue?.ToString();
            //    double dBankCost = 0;
            //    if (!string.IsNullOrEmpty(sBankCost)) dBankCost = Convert.ToDouble(sBankCost);

            //    string sSurvey = TxtSurvey.EditValue?.ToString();
            //    double dSurvey = 0;
            //    if (!string.IsNullOrEmpty(sSurvey)) dSurvey = Convert.ToDouble(sSurvey);

            //    string sTotAmt = TxtTotAmt.EditValue?.ToString();
            //    double dTotAmt = 0;
            //    if (!string.IsNullOrEmpty(sTotAmt)) dTotAmt = Convert.ToDouble(sTotAmt);

            //    string sAvgUnpr = TxtAvgUnpr.EditValue?.ToString();
            //    double dAvgUnpr = 0;
            //    if (!string.IsNullOrEmpty(sAvgUnpr)) dAvgUnpr = Convert.ToDouble(sAvgUnpr);

            //    string sIcmDealerNm = BtnEditIcmCompNm.EditValue?.ToString();

            //    string sIcmWeight = TxtWareHWeight.EditValue?.ToString();
            //    double dIcmWeight = 0;
            //    if (!string.IsNullOrEmpty(sIcmWeight)) dIcmWeight = Convert.ToDouble(sIcmWeight);
                
            //    string sStartDirectory = Application.StartupPath + "\\ExcelForm";
            //    string sFileName = "ImportScrapDailyReportFormat.xlsx";
                
            //    string sCopyDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //    string sCopyFileName = "수입고철일보" + sBlNo + ".xlsx";

            //    File.Copy(Path.Combine(sStartDirectory, sFileName), Path.Combine(sCopyDirectory, sCopyFileName), true);

            //    excelApp = new Excel.Application();
            //    wb = excelApp.Workbooks.Open(sCopyDirectory + "\\" + sCopyFileName, 0, false);
            //    ws = wb.Sheets["수입고철일보"];

            //    if (!string.IsNullOrEmpty(sPmntMtd))
            //    {
            //        ws.Cells[6, 2] = "T/T";
            //    }
            //    else
            //    {
            //        ws.Cells[6, 2] = sLcNo;         //L/C No
            //    }
            //    ws.Cells[6, 7] = sGrade;            //Grade
            //    ws.Cells[6, 10] = sArrvDate;        //ARRV_DATE
            //    ws.Cells[6, 12] = sTransOprtnComp;  //TRANS_OPRTN_COMPANY
            //    ws.Cells[7, 2] = sBlNo;             //B/L No
            //    ws.Cells[7, 7] = dContTon;          //CONT_TON
            //    ws.Cells[7, 10] = dContAmt;         //CONT_AMT
            //    ws.Cells[7, 12] = dContTotAmt;      //CONT_TOT_AMT
            //    ws.Cells[8, 2] = dImtTon;           //ICM_TON
            //    ws.Cells[8, 7] = dExchangeRate;     //ICM_EXCHANGE_RATE
            //    ws.Cells[8, 10] = dIcmAmt;          //ICM_AMT
            //    ws.Cells[9, 2] = dBrokerCost;       //BROKER_COST
            //    ws.Cells[9, 7] = dBankCost;         //BANK_COST
            //    ws.Cells[9, 11] = dSurvey;          //SURVEY
            //    ws.Cells[10, 2] = dTotAmt;          //TOT_AMT
            //    ws.Cells[10, 7] = dAvgUnpr;         //AVG_UNPR
            //    ws.Cells[10, 10] = sIcmDealerNm;    //ICM_DEALER_NM
            //    ws.Cells[10, 13] = dIcmWeight;      //ICM_WEIGHT

            //    for (int i = 0; i < GridViewDtlRetr.RowCount; i++) //스크랩 중량
            //    {
            //        DataRow drDtl = GridViewDtlRetr.GetDataRow(i);

            //        ws.Cells[12 + i, 1] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_DEALER_CD");
            //        ws.Cells[12 + i, 2] = GridViewDtlRetr.GetRowCellDisplayText(i, "VEHICLE_NO");
            //        ws.Cells[12 + i, 3] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_WEIGHT");
            //        ws.Cells[12 + i, 4] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_UNPR");
            //        ws.Cells[12 + i, 5] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_GRADE");
            //        ws.Cells[12 + i, 6] = GridViewDtlRetr.GetRowCellDisplayText(i, "IN_WAREHOUSE_AMT");
            //        ws.Cells[12 + i, 7] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_PAYMENT");
            //        ws.Cells[12 + i, 8] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_LOSS_WEIGHT");
            //        ws.Cells[12 + i, 9] = GridViewDtlRetr.GetRowCellDisplayText(i, "ICM_CARRY_COST");
            //    }
            //    Cursor = Cursors.Default;

            //    wb.Save();

            //    ws = null;
            //    wb = null;
            //    excelApp.Quit();

            //    Process.Start(sCopyDirectory + "\\" + sCopyFileName);
            //}
            //catch (Exception ex)
            //{
            //    Cursor = Cursors.WaitCursor;
            //    XtraMessageBox.Show(ex.ToString());
            //    Cursor = Cursors.Default;

            //    ws = null;
            //    wb = null;
            //    if(excelApp != null)
            //        excelApp.Quit();
            //}
        }

        private void BtnEditGradeNm_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpGradeCd frm = new PopUpGradeCd();

            DataRow drGradeInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drGradeInfo = frm.drResult;
                BtnEditGradeNm.EditValue = drGradeInfo["GUBUN1"].ToString();
                TxtGradeCd.EditValue = drGradeInfo["J_SERIAL"];
            }
        }

        private void BtnEditShipCompNm_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpDealerCd frm = new PopUpDealerCd();

            DataRow drDealerInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drDealerInfo = frm.drResult;
                BtnEditShipCompNm.EditValue = drDealerInfo["DEALER_NM"].ToString();
                TxtShipCompCd.EditValue = drDealerInfo["DEALER_CD"].ToString();
            }
        }

        private void BtnEditIcmCompNm_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpDealerCd frm = new PopUpDealerCd();

            DataRow drDealerInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drDealerInfo = frm.drResult;
                BtnEditIcmCompNm.EditValue = drDealerInfo["DEALER_NM"].ToString();
                TxtIcmCompCd.EditValue = drDealerInfo["DEALER_CD"].ToString();
            }
        }

        private void TxtContAmt_Leave(object sender, EventArgs e)
        {
            string sContAmt = TxtContAmt.EditValue?.ToString();
            string sContTon = TxtContTon.EditValue?.ToString();

            if(!string.IsNullOrEmpty(sContAmt) && !string.IsNullOrEmpty(sContTon))
            {
                TxtContTotAmt.EditValue = Convert.ToDouble(sContAmt) * Convert.ToDouble(sContTon);
            }

            string sIcmWeight = TxtIncomeTon.EditValue?.ToString();
            string sIcmTotAmt = TxtIcmAmt.EditValue?.ToString();

            if(!string.IsNullOrEmpty(sContAmt) && !string.IsNullOrEmpty(sIcmWeight) && !string.IsNullOrEmpty(sIcmTotAmt))
            {
                TxtExchangeRate.EditValue = Convert.ToDouble(sIcmTotAmt) / (Convert.ToDouble(sIcmWeight) * Convert.ToDouble(sContAmt));
            }
        }

        private void TxtContTon_Leave(object sender, EventArgs e)
        {
            string sContAmt = TxtContAmt.EditValue?.ToString();
            string sContTon = TxtContTon.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sContAmt) && !string.IsNullOrEmpty(sContTon))
            {
                TxtContTotAmt.EditValue = Convert.ToDouble(sContAmt) * Convert.ToDouble(sContTon);
            }
        }

        private void TxtIncomeTon_Leave(object sender, EventArgs e)
        {
            string sContAmt = TxtContAmt.EditValue?.ToString();
            string sIcmWeight = TxtIncomeTon.EditValue?.ToString();
            string sIcmTotAmt = TxtIcmAmt.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sContAmt) && !string.IsNullOrEmpty(sIcmWeight) && !string.IsNullOrEmpty(sIcmTotAmt))
            {
                TxtExchangeRate.EditValue = Convert.ToDouble(sIcmTotAmt) / (Convert.ToDouble(sIcmWeight) * Convert.ToDouble(sContAmt));
            }
        }

        private void TxtIcmAmt_Leave(object sender, EventArgs e)
        {
            string sContAmt = TxtContAmt.EditValue?.ToString();
            string sIcmWeight = TxtIncomeTon.EditValue?.ToString();
            string sIcmTotAmt = TxtIcmAmt.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sContAmt) && !string.IsNullOrEmpty(sIcmWeight) && !string.IsNullOrEmpty(sIcmTotAmt))
            {
                TxtExchangeRate.EditValue = Convert.ToDouble(sIcmTotAmt) / (Convert.ToDouble(sIcmWeight) * Convert.ToDouble(sContAmt));
            }

            double dIcmTotAMt = 0;
            if (!string.IsNullOrEmpty(sIcmTotAmt)) dIcmTotAMt = Convert.ToDouble(TxtIcmAmt.EditValue);
            
            string sBrokerCost = TxtBrokerCost.EditValue?.ToString();
            double dBrokerCost = 0;
            if (!string.IsNullOrEmpty(sBrokerCost)) dBrokerCost = Convert.ToDouble(TxtBrokerCost.EditValue);


            string sBankCost = TxtBankCost.EditValue?.ToString();
            double dBankCost = 0;
            if (!string.IsNullOrEmpty(sBankCost)) dBankCost = Convert.ToDouble(TxtBankCost.EditValue);

            string sSurvey = TxtSurvey.EditValue?.ToString();
            double dSurvey = 0;
            if (!string.IsNullOrEmpty(sSurvey)) dSurvey = Convert.ToDouble(TxtSurvey.EditValue);

            TxtTotAmt.EditValue = dIcmTotAMt + dBrokerCost + dBankCost + dSurvey;
        }

        private void TxtBrokerCost_Leave(object sender, EventArgs e)
        {
            string sIcmTotAmt = TxtIcmAmt.EditValue?.ToString();
            double dIcmTotAMt = 0;
            if (!string.IsNullOrEmpty(sIcmTotAmt)) dIcmTotAMt = Convert.ToDouble(TxtIcmAmt.EditValue);

            string sBrokerCost = TxtBrokerCost.EditValue?.ToString();
            double dBrokerCost = 0;
            if (!string.IsNullOrEmpty(sBrokerCost)) dBrokerCost = Convert.ToDouble(TxtBrokerCost.EditValue);
            
            string sBankCost = TxtBankCost.EditValue?.ToString();
            double dBankCost = 0;
            if (!string.IsNullOrEmpty(sBankCost)) dBankCost = Convert.ToDouble(TxtBankCost.EditValue);

            string sSurvey = TxtSurvey.EditValue?.ToString();
            double dSurvey = 0;
            if (!string.IsNullOrEmpty(sSurvey)) dSurvey = Convert.ToDouble(TxtSurvey.EditValue);

            TxtTotAmt.EditValue = dIcmTotAMt + dBrokerCost + dBankCost + dSurvey;
        }

        private void TxtBankCost_Leave(object sender, EventArgs e)
        {
            string sIcmTotAmt = TxtIcmAmt.EditValue?.ToString();
            double dIcmTotAMt = 0;
            if (!string.IsNullOrEmpty(sIcmTotAmt)) dIcmTotAMt = Convert.ToDouble(TxtIcmAmt.EditValue);

            string sBrokerCost = TxtBrokerCost.EditValue?.ToString();
            double dBrokerCost = 0;
            if (!string.IsNullOrEmpty(sBrokerCost)) dBrokerCost = Convert.ToDouble(TxtBrokerCost.EditValue);
            
            string sBankCost = TxtBankCost.EditValue?.ToString();
            double dBankCost = 0;
            if (!string.IsNullOrEmpty(sBankCost)) dBankCost = Convert.ToDouble(TxtBankCost.EditValue);

            string sSurvey = TxtSurvey.EditValue?.ToString();
            double dSurvey = 0;
            if (!string.IsNullOrEmpty(sSurvey)) dSurvey = Convert.ToDouble(TxtSurvey.EditValue);

            TxtTotAmt.EditValue = dIcmTotAMt + dBrokerCost + dBankCost + dSurvey;
        }

        private void TxtSurvey_Leave(object sender, EventArgs e)
        {
            string sIcmTotAmt = TxtIcmAmt.EditValue?.ToString();
            double dIcmTotAMt = 0;
            if (!string.IsNullOrEmpty(sIcmTotAmt)) dIcmTotAMt = Convert.ToDouble(TxtIcmAmt.EditValue);

            string sBrokerCost = TxtBrokerCost.EditValue?.ToString();
            double dBrokerCost = 0;
            if (!string.IsNullOrEmpty(sBrokerCost)) dBrokerCost = Convert.ToDouble(TxtBrokerCost.EditValue);
            
            string sBankCost = TxtBankCost.EditValue?.ToString();
            double dBankCost = 0;
            if (!string.IsNullOrEmpty(sBankCost)) dBankCost = Convert.ToDouble(TxtBankCost.EditValue);

            string sSurvey = TxtSurvey.EditValue?.ToString();
            double dSurvey = 0;
            if (!string.IsNullOrEmpty(sSurvey)) dSurvey = Convert.ToDouble(TxtSurvey.EditValue);

            TxtTotAmt.EditValue = dIcmTotAMt + dBrokerCost + dBankCost + dSurvey;
        }

        private void GridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("ICM_CONT_DATE") || e.Column.FieldName.Equals("ARRV_DATE"))
            {
                if (e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
        }

        private void LkupPmntMethod_Leave(object sender, EventArgs e)
        {
            string sChkPmntMtd = LkupPmntMethod.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sChkPmntMtd))
            {
                if(sChkPmntMtd == "1")
                {
                    TxtLcNo.EditValue = string.Empty;
                    TxtLcNo.ReadOnly = true;
                }
                else
                {
                    TxtLcNo.ReadOnly = false;
                }
            }
            else
            {
                TxtLcNo.ReadOnly = false;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void IncomeScrapDailyReport_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnNew_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
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

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridDtlRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void IncomeScrapDailyReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void IncomeScrapDailyReport_TextChanged(object sender, EventArgs e)
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

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}