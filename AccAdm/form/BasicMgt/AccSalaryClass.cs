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
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.EditForm.Helpers.Controls;
using System.Diagnostics;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class AccSalaryClass : DevExpress.XtraEditors.XtraForm
    {
        public AccSalaryClass()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        private void AccSalaryClass_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DtpDate.EditValue = DateTime.Now;
            DataTable dtJobType = GetLookUpData("1", "Y", "");
            ComLib.ComGrid.SetLookUpEdit(LkupEditJobType, dtJobType, "CD", "NM", "Y");

            RepositoryItemGridLookUpEdit jobcLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(jobcLkup, dtJobType, GridRight, GridColAlwRefrCd1, "CD", "NM", "");

            DataTable dtPmntRefr = GetLookUpData("2", "N", "");
            RepositoryItemGridLookUpEdit pmntRefrLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(pmntRefrLkup, dtPmntRefr, GridRight, GridColPmntRefr, "CD", "NM", "");

            DataTable dtPmntGb = GetLookUpData("3", "N", "");
            RepositoryItemGridLookUpEdit pmntGbLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(pmntGbLkup, dtPmntGb, GridRight, GridColPmntGb, "CD", "NM", "");
            DataTable dtGrade = GetLookUpData("4", "N", "");
            RepositoryItemGridLookUpEdit gradeLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(gradeLkup, dtGrade, GridLeft, GridColLAlwRefrCd2, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(gradeLkup, dtGrade, GridRight, GridColAlwRefrCd2, "CD", "NM", "");
            DataTable dtPayCd = GetLookUpData("5", "N", "");
            RepositoryItemGridLookUpEdit payCdLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(payCdLkup, dtPayCd, GridRight, GridColPayCd, "CD", "NM", "");

            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccSalaryClass");

            ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccSalaryClass", GridViewLeft);
            ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccSalaryClass1", GridViewRight);
        }
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                //  strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine("SELECT A.COM_CD AS CD ");
                strSql.AppendLine("     , A.COM_NM AS NM  ");
                strSql.AppendLine("  FROM COM_BASE_CD A ");
                strSql.AppendLine(" WHERE CD_GB='ALW_REFR_CD1'");


            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine("SELECT A.COM_CD AS CD ");
                strSql.AppendLine("     , A.COM_NM AS NM  ");
                strSql.AppendLine("  FROM COM_BASE_CD A ");
                strSql.AppendLine(" WHERE CD_GB='PMNT_REFR'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine("SELECT A.COM_CD AS CD ");
                strSql.AppendLine("     , A.COM_NM AS NM  ");
                strSql.AppendLine("  FROM COM_BASE_CD A ");
                strSql.AppendLine(" WHERE CD_GB='PMNT_GB'");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine("SELECT A.COM_CD AS CD ");
                strSql.AppendLine("     , A.COM_NM AS NM  ");
                strSql.AppendLine("  FROM COM_BASE_CD A ");
                strSql.AppendLine(" WHERE CD_GB='GRADE_CD'");
            }
            else if (sGb.Equals("5"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD       ");
                strSql.AppendLine("      , COM_NM AS NM       ");
                strSql.AppendLine("   FROM COM_BASE_CD        ");
                strSql.AppendLine("  WHERE CD_GB = 'ACPAYDTL1'");
            }
            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }
        private void GetLeft()
        {
            StringBuilder strSql = new StringBuilder();
            string sAlwRefrCd1 = LkupEditJobType.EditValue.ToString();         
            strSql.Clear();
            strSql.AppendLine(" SELECT A.ALW_REFR_CD1            ");
            strSql.AppendLine(" 	 , B.COM_NM AS COM_NM         ");
            strSql.AppendLine(" 	 , ALW_REFR_CD2               ");
            strSql.AppendLine("	  FROM ACC_PAY_ALW_REFR A    ");
            strSql.AppendLine("   LEFT OUTER JOIN    COM_BASE_CD B                        ");
            strSql.AppendLine("     ON A.ALW_REFR_CD1 = B.COM_CD AND B.CD_GB='ALW_REFR_CD1'   ");
            strSql.AppendLine("  WHERE 1=1     ");
            strSql.AppendLine("    AND (('****' = '" + sAlwRefrCd1 + "') OR (('****' <> '" + sAlwRefrCd1 + "') AND A.ALW_REFR_CD1 = '" + sAlwRefrCd1 + "'))");              
            strSql.AppendLine("  GROUP BY ALW_REFR_CD1,ALW_REFR_CD2,COM_NM, SORT_SEQ     ");
            strSql.AppendLine("  ORDER BY A.ALW_REFR_CD1 ,B.SORT_SEQ    ");            
            DataTable dtL = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridLeft.DataSource = dtL;
        }
        private void GetRight()
        {
            StringBuilder strSql = new StringBuilder();
            string sAlwRefrCd1 = GridViewLeft.GetFocusedRowCellValue("ALW_REFR_CD1").ToString();
            string sAlwRefrCd2 = GridViewLeft.GetFocusedRowCellValue("ALW_REFR_CD2").ToString();
            string sDate = DtpDate.EditValue.ToString().Replace("-", "").Substring(0, 6);
            strSql.AppendLine("");
            strSql.AppendLine("SELECT A.PAY_CD ");
            strSql.AppendLine("     , ALW_REFR_CD1");//직종명
            strSql.AppendLine("     , A.ALW_REFR_CD2");
            strSql.AppendLine("     , A.ALW_REFR_CD3");
            strSql.AppendLine("     , A.STRT_MM");
            strSql.AppendLine("     , A.END_MM");
            strSql.AppendLine("     , A.PMNT_REFR");//지급기준
            strSql.AppendLine("     , A.PMNT_GB ");//지급구분
            strSql.AppendLine("     , A.PMNT_RATE");
            strSql.AppendLine("     , A.PMNT_AMT");
            strSql.AppendLine("     , A.ENT_DT");
            strSql.AppendLine("     , A.ENT_ID");
            strSql.AppendLine("     , A.MFY_ID");
            strSql.AppendLine("     , A.MFY_DT");
            strSql.AppendLine("  FROM ACC_PAY_ALW_REFR A");
            strSql.AppendLine(" WHERE A.ALW_REFR_CD1 = '" + sAlwRefrCd1 + "'");
            strSql.AppendLine("   AND A.ALW_REFR_CD2  = '" + sAlwRefrCd2 + "'");
            strSql.AppendLine("   AND '" + sDate + "' >= STRT_MM   "); 
            strSql.AppendLine("   AND '" + sDate + "' <  END_MM    ");
            strSql.AppendLine(" ORDER BY LEN(ALW_REFR_CD3),ALW_REFR_CD3 ");

            DataTable dtR = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRight.DataSource = dtR;
        }

        private void GridViewLeft_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            GetRight();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetLeft();
            GetRight();
        }
        private void LkupEditJobType_EditValueChanged(object sender, EventArgs e)
        {
            GetLeft();
            GetRight();
        }
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sALW_REFR_CD1 = GridViewLeft.GetFocusedRowCellValue("ALW_REFR_CD1")?.ToString();
            string sALW_REFR_CD2 = GridViewLeft.GetFocusedRowCellValue("ALW_REFR_CD2")?.ToString();

            if(GridViewRight.RowCount > 0)
            {
                string sPAY_CD = GridViewRight.GetRowCellValue(GridViewRight.RowCount - 1, "PAY_CD")?.ToString();
                string sALW_REFR_CD3 = GridViewRight.GetRowCellValue(GridViewRight.RowCount - 1, "ALW_REFR_CD3")?.ToString();

                if (string.IsNullOrEmpty(sPAY_CD))
                {
                    XtraMessageBox.Show("급여코드를 입력해주세요.");
                    return;
                }
                else if (string.IsNullOrEmpty(sALW_REFR_CD3))
                {
                    XtraMessageBox.Show("호봉을 입력해주세요.");
                    return;
                }
            }

            GridViewRight.AddNewRow();
            GridViewRight.Focus();
            GridViewRight.SetFocusedRowCellValue("ALW_REFR_CD1", sALW_REFR_CD1);
            GridViewRight.SetFocusedRowCellValue("ALW_REFR_CD2", sALW_REFR_CD2);
        }
        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if ((MessageBox.Show(this, "저장 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
            {
                return;
            }
            Cursor = Cursors.WaitCursor;

            DataTable dt = (DataTable)GridRight.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            StringBuilder strSql = new StringBuilder();
            if (dtMerge.Rows.Count > 0)  // modify
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    string sPayCd = string.Empty;
                    string sAlwRefrCd1 = string.Empty;
                    string sAlwRefrCd2 = string.Empty;
                    string sAlwRefrCd3 = string.Empty;
                    string sStrtMm = string.Empty;
                    string sEndMm = string.Empty;
                    string sPmntRefr = string.Empty;
                    string sPmntGb = string.Empty;
                    double sPmntRate = 0;
                    double sPmntAmt = 0;
                    string sEntDt = string.Empty;
                    string sId = FmMainToolBar2.drUser["USRCD"]?.ToString();

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        sPayCd = Convert.ToString(dtMerge.Rows[j]["PAY_CD"]);
                        sAlwRefrCd1 = Convert.ToString(dtMerge.Rows[j]["ALW_REFR_CD1"]);
                        sAlwRefrCd2 = Convert.ToString(dtMerge.Rows[j]["ALW_REFR_CD2"]);
                        sAlwRefrCd3 = Convert.ToString(dtMerge.Rows[j]["ALW_REFR_CD3"]);
                        sStrtMm = Convert.ToString(dtMerge.Rows[j]["STRT_MM"]);
                        sEndMm = Convert.ToString(dtMerge.Rows[j]["END_MM"]);
                        sPmntRefr = Convert.ToString(dtMerge.Rows[j]["PMNT_REFR"]);
                        sPmntGb = Convert.ToString(dtMerge.Rows[j]["PMNT_GB"]);
                        double.TryParse(dtMerge.Rows[j]["PMNT_RATE"]?.ToString(), out sPmntRate);
                        double.TryParse(dtMerge.Rows[j]["PMNT_AMT"]?.ToString(), out sPmntAmt);

                        if (string.IsNullOrEmpty(sPayCd))
                        {
                            MessageBox.Show("급여코드를 입력하세요");
                            GridViewRight.FocusedRowHandle = GridViewRight.LocateByDisplayText(0, GridColPayCd, sPayCd);
                            GridViewRight.FocusedColumn = GridColPayCd;
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }
                        if (string.IsNullOrEmpty(sAlwRefrCd1))
                        {
                            MessageBox.Show("직종을  입력하세요");
                            GridViewRight.FocusedRowHandle = GridViewRight.LocateByDisplayText(0, GridColAlwRefrCd1, sAlwRefrCd1);
                            GridViewRight.FocusedColumn = GridColAlwRefrCd1;
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }
                        if (string.IsNullOrEmpty(sAlwRefrCd2))
                        {
                            MessageBox.Show("직급을  입력하세요");
                            GridViewRight.FocusedRowHandle = GridViewRight.LocateByDisplayText(0, GridColAlwRefrCd2, sAlwRefrCd2);
                            GridViewRight.FocusedColumn = GridColAlwRefrCd2;
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }
                        if (string.IsNullOrEmpty(sAlwRefrCd3))
                        {
                            MessageBox.Show("호봉을 입력하세요");
                            GridViewRight.FocusedRowHandle = GridViewRight.LocateByDisplayText(0, GridColAlwRefrCd3, sAlwRefrCd3);
                            GridViewRight.FocusedColumn = GridColAlwRefrCd3;
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }
                        if (string.IsNullOrEmpty(sStrtMm))
                        {
                            MessageBox.Show("시작월을 입력하세요");
                            GridViewRight.FocusedRowHandle = GridViewRight.LocateByDisplayText(0, GridColStrtMm, sStrtMm);
                            GridViewRight.FocusedColumn = GridColStrtMm;
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }

                        strSql.Clear();
                        strSql.AppendLine(" MERGE INTO ACC_PAY_ALW_REFR AS a                                                                                                                                 ");
                        strSql.AppendLine("     USING(SELECT                                                                                                                                                 ");
                        strSql.AppendLine("                  PAY_CD = '" + sPayCd + "'                                                                                                                                              ");
                        strSql.AppendLine("                , ALW_REFR_CD1 = '" + sAlwRefrCd1 + "'                                                                                                                               ");
                        strSql.AppendLine("                , ALW_REFR_CD2 = '" + sAlwRefrCd2 + "'                                                                                                                               ");
                        strSql.AppendLine("                , ALW_REFR_CD3 = '" + sAlwRefrCd3 + "'                                                                                                                               ");
                        strSql.AppendLine("                , STRT_MM = '" + sStrtMm + "'                                                                                                                                    ");
                        strSql.AppendLine("         , END_MM = '" + sEndMm + "'                                                                                                                                      ");
                        strSql.AppendLine("                , PMNT_REFR = '" + sPmntRefr + "'                                                                                                                                 ");
                        strSql.AppendLine("                , PMNT_GB = '" + sPmntGb + "'                                                                                                                                   ");
                        strSql.AppendLine("                , PMNT_RATE = '" + sPmntRate + "'                                                                                                                                 ");
                        strSql.AppendLine("                , PMNT_AMT = '" + sPmntAmt + " '                                                                                                                              ");
                        strSql.AppendLine("                , ENT_DT = CONVERT(VARCHAR(20),GETDATE(),20)                                                                                                                              ");
                        strSql.AppendLine("                , ENT_ID = '"+ sId + "') AS b                                                                                                                         ");
                        strSql.AppendLine("     ON a.PAY_CD = b.PAY_CD AND a.ALW_REFR_CD1 = b.ALW_REFR_CD1 AND a.ALW_REFR_CD2 = b.ALW_REFR_CD2 AND a.ALW_REFR_CD3 = b.ALW_REFR_CD3 AND a.STRT_MM = b.STRT_MM ");
                        strSql.AppendLine("     WHEN MATCHED THEN UPDATE SET                                                                                                                                 ");
                        strSql.AppendLine("         PAY_CD = '" + sPayCd + "'                                                                                                                                              ");
                        strSql.AppendLine("                , ALW_REFR_CD1 = '" + sAlwRefrCd1 + "'                                                                                                                               ");
                        strSql.AppendLine("                , ALW_REFR_CD2 = '" + sAlwRefrCd2 + "'                                                                                                                               ");
                        strSql.AppendLine("                , ALW_REFR_CD3 = '" + sAlwRefrCd3 + "'                                                                                                                               ");
                        strSql.AppendLine("                , STRT_MM = '" + sStrtMm + "'                                                                                                                                    ");
                        strSql.AppendLine("         , END_MM = '" + sEndMm + "'                                                                                                                                      ");
                        strSql.AppendLine("                , PMNT_REFR = '" + sPmntRefr + "'                                                                                                                                 ");
                        strSql.AppendLine("                , PMNT_GB = '" + sPmntGb + "'                                                                                                                                   ");
                        strSql.AppendLine("                , PMNT_RATE = '" + sPmntRate + "'                                                                                                                                 ");
                        strSql.AppendLine("                , PMNT_AMT = '" + sPmntAmt + " '                                                                                                                              ");
                        strSql.AppendLine("                , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20)                                                                                                                              ");
                        strSql.AppendLine("                , MFY_ID = '"+ sId + "'                                                                                                                               ");
                        strSql.AppendLine("     WHEN NOT MATCHED THEN INSERT(                                                                                                                                ");
                        strSql.AppendLine("         PAY_CD                                                                                                                                                   ");
                        strSql.AppendLine("                , ALW_REFR_CD1                                                                                                                                    ");
                        strSql.AppendLine("                , ALW_REFR_CD2                                                                                                                                    ");
                        strSql.AppendLine("                , ALW_REFR_CD3                                                                                                                                    ");
                        strSql.AppendLine("                , STRT_MM                                                                                                                                         ");
                        strSql.AppendLine("                , END_MM                                                                                                                                          ");
                        strSql.AppendLine("                , PMNT_REFR                                                                                                                                       ");
                        strSql.AppendLine("                , PMNT_GB                                                                                                                                         ");
                        strSql.AppendLine("                , PMNT_RATE                                                                                                                                       ");
                        strSql.AppendLine("                , PMNT_AMT                                                                                                                                        ");
                        strSql.AppendLine("                , ENT_DT                                                                                                                                          ");
                        strSql.AppendLine("                , ENT_ID)                                                                                                                                         ");
                        strSql.AppendLine("     VALUES(                                                                                                                                                      ");
                        strSql.AppendLine("                  '" + sPayCd + "'                                                                                                                                                   ");
                        strSql.AppendLine("                , '" + sAlwRefrCd1 + "'                                                                                                                                              ");
                        strSql.AppendLine("                , '" + sAlwRefrCd2 + "'                                                                                                                                              ");
                        strSql.AppendLine("                , '" + sAlwRefrCd3 + "'                                                                                                                                              ");
                        strSql.AppendLine("                , '" + sStrtMm + "'                                                                                                                                              ");
                        strSql.AppendLine("                , '" + sEndMm + "'                                                                                                                                        ");
                        strSql.AppendLine("                , '" + sPmntRefr + "'                                                                                                                                             ");
                        strSql.AppendLine("                , '" + sPmntGb + "'                                                                                                                                             ");
                        strSql.AppendLine("                , '" + sPmntRate + "'                                                                                                                                             ");
                        strSql.AppendLine("                , '" + sPmntAmt + " '                                                                                                                                         ");
                        strSql.AppendLine("                , CONVERT(VARCHAR(20),GETDATE(),20)                                                                                                                                       ");
                        strSql.AppendLine("                , '"+ sId + "'                                                                                                                                    ");
                        strSql.AppendLine("     );                                                                                                                                                           ");

                        /*
                        strSql.AppendLine("INSERT INTO ACC_PAY_ALW_REFR ");
                        strSql.AppendLine("           (PAY_CD       ");
                        strSql.AppendLine("           ,ALW_REFR_CD1 ");
                        strSql.AppendLine("           ,ALW_REFR_CD2 ");
                        strSql.AppendLine("           ,ALW_REFR_CD3 ");
                        strSql.AppendLine("           ,STRT_MM      ");
                        strSql.AppendLine("           ,END_MM       ");
                        strSql.AppendLine("           ,PMNT_REFR    ");
                        strSql.AppendLine("           ,PMNT_GB      ");
                        strSql.AppendLine("           ,PMNT_RATE    ");
                        strSql.AppendLine(" 	      ,PMNT_AMT     ");
                        strSql.AppendLine(" 	      ,ENT_DT      ");
                        strSql.AppendLine(" 	      ,ENT_ID       ");
                        strSql.AppendLine(" 	      ,MFY_DT       ");
                        strSql.AppendLine(" 	      ,MFY_ID )     ");
                        strSql.AppendLine("     VALUES (");
                        strSql.AppendLine("            '" + sPayCd + "'");
                        strSql.AppendLine(" 		  ,'" + sAlwRefrCd1 + "'   ");
                        strSql.AppendLine(" 		  ,'" + sAlwRefrCd2 + "'   ");
                        strSql.AppendLine(" 		  ,'" + sAlwRefrCd3 + "'   ");
                        strSql.AppendLine(" 		  ,'" + sStrtMm + "'    ");
                        strSql.AppendLine(" 		  ,'" + sEndMm + "'     ");
                        strSql.AppendLine(" 		  ,'" + sPmntRefr + "'  ");
                        strSql.AppendLine(" 		  ,'" + sPmntGb + "'    ");
                        strSql.AppendLine(" 		  ,'" + sPmntRate + "'  ");
                        strSql.AppendLine(" 		  ,'" + sPmntAmt + "'   ");              
                        strSql.AppendLine(" 		  ,NOW()    ");
                        strSql.AppendLine(" 		  , '입력ID'   ");
                        strSql.AppendLine(" 		  ,NOW()     ");
                        strSql.AppendLine(" 		  , '입력ID동일'  ");
                        strSql.AppendLine("            )                    ");
                        strSql.AppendLine("         ON DUPLICATE KEY UPDATE              ");
                        strSql.AppendLine("  		   END_MM      ='" + sEndMm + "'     ");
                        strSql.AppendLine("           ,PMNT_REFR   ='" + sPmntRefr + "'  ");
                        strSql.AppendLine(" 	      ,PMNT_GB     ='" + sPmntGb + "'    ");
                        strSql.AppendLine(" 	      ,PMNT_RATE   ='" + sPmntRate + "'  ");
                        strSql.AppendLine(" 	      ,PMNT_AMT    ='" + sPmntAmt + "'   ");
                        strSql.AppendLine(" 	      ,MFY_DT      = NOW()    ");
                        strSql.AppendLine(" 	      ,MFY_ID      ='수정ID'    ");
                        */

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "Table:ACC_PAY_ALW_REFR -> PAY_CD:" + sPayCd + ",ALW_REFR_CD1:" + sAlwRefrCd1 + ",ALW_REFR_CD2:" + sAlwRefrCd2 + ",ALW_REFR_CD3:" + sAlwRefrCd3 + ",STRT_MM:" + sStrtMm;
                        ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, (j+1).ToString(), "S", this.Name, sLogRmk, cmd);
                    }                
                    
                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");       
                    
                    GetRight();
                    GridViewRight.FocusedRowHandle = GridViewRight.LocateByDisplayText(0, GridColAlwRefrCd3, dtMerge.Rows[0]["ALW_REFR_CD3"]?.ToString());


                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                XtraMessageBox.Show("변경된 데이터가 없습니다.");
            }
            Cursor = Cursors.Default;

        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sPayCd = GridViewRight.GetFocusedRowCellValue("PAY_CD")?.ToString();
            string sAlwRefrCd1 = GridViewRight.GetFocusedRowCellValue("ALW_REFR_CD1")?.ToString();
            string sAlwRefrCd2 = GridViewRight.GetFocusedRowCellValue("ALW_REFR_CD2")?.ToString();
            string sAlwRefrCd3 = GridViewRight.GetFocusedRowCellValue("ALW_REFR_CD3")?.ToString();
            string sStrtMM = GridViewRight.GetFocusedRowCellValue("STRT_MM")?.ToString();

            if (XtraMessageBox.Show("급여코드 : " + sPayCd + "\r\n직종코드 : " + sAlwRefrCd1 
                + " \r\n직급코드 : " + sAlwRefrCd2 + "\r\n 호봉코드 : " + sAlwRefrCd3 
                + "\r\n 시작월 : " + sStrtMM + "\r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
               , "호봉현황 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }
            
            if(string.IsNullOrEmpty(sPayCd) || string.IsNullOrEmpty(sAlwRefrCd1) || string.IsNullOrEmpty(sAlwRefrCd2) 
                || string.IsNullOrEmpty(sAlwRefrCd3) || string.IsNullOrEmpty(sStrtMM))
            {
                XtraMessageBox.Show("필수 삭제 정보가 존재하지 않습니다.");
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM ACC_PAY_ALW_REFR ");
                strSql.AppendLine("       WHERE PAY_CD = '" + sPayCd + "' ");
                strSql.AppendLine("         AND ALW_REFR_CD1 = '" + sAlwRefrCd1 + "' ");
                strSql.AppendLine("         AND ALW_REFR_CD2 = '" + sAlwRefrCd2 + "' ");
                strSql.AppendLine("         AND ALW_REFR_CD3 = '" + sAlwRefrCd3 + "' ");
                strSql.AppendLine("         AND STRT_MM = '" + sStrtMM + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:ACC_PAY_ALW_REFR -> PAY_CD:" + sPayCd + ",ALW_REFR_CD1:" + sAlwRefrCd1 + ",ALW_REFR_CD2:" + sAlwRefrCd2 + ",ALW_REFR_CD3:" + sAlwRefrCd3 + ",STRT_MM:" + sStrtMM;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRight.FocusedRowHandle;
                GetRight();
                GridViewRight.FocusedRowHandle = idx - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void GridViewRight_ShowingPopupEditForm(object sender, DevExpress.XtraGrid.Views.Grid.ShowingPopupEditFormEventArgs e)
        {
            e.EditForm.ImeMode = ImeMode.Hangul;
            e.EditForm.StartPosition = FormStartPosition.CenterParent;

            foreach (var button in e.EditForm.Controls.OfType<EditFormContainer>()
                        .SelectMany(control => Enumerable.Cast<Control>(control.Controls)).OfType<PanelControl>()
                        .SelectMany(nestedControl => Enumerable.Cast<Control>(nestedControl.Controls)).OfType<SimpleButton>())
            {
                switch (button.Text)
                {
                    case "Update":
                        //button.Visible = false;
                        button.Text = "수정";
                        //button.Click += EditFormUpdateButton_Click;
                        break;

                    case "Cancel":
                        button.Text = "취소";
                        break;
                }
            }

        }

        private void AccSalaryClass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
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

        private void GridLeft_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridRight_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewLeft_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccSalaryClass_FormClosed(object sender, FormClosedEventArgs e)
        {
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccSalaryClass", GridViewLeft);
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccSalaryClass1", GridViewRight);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = "봉급표리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridRight.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
                }
                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        private void GridViewRight_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRight.UpdateCurrentRow();
        }

        private void DtpDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}