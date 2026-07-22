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
using DevExpress.XtraEditors.Repository;
using ComLib;
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.EditForm.Helpers.Controls;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class AccCreditCardMgt : DevExpress.XtraEditors.XtraForm
    {
        public AccCreditCardMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccCreditCardMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            DataTable dtCardGb = GetLookUpData("1", "1", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditCardGB, dtCardGb, "CD", "NM", "Y");
            DataTable dtCardCompany = GetLookUpData("2", "1", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditCardCompany, dtCardCompany, "CD", "NM", "Y");
            LkupEditCardGB.EditValue = "";
            LkupEditCardCompany.EditValue = "";
            RdGrUseYn.EditValue = "Y";

            DataTable dtCardKind = GetLookUpData("3", "1", "", "Y");

            RepositoryItemGridLookUpEdit cardGbLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(cardGbLkup, dtCardGb, GridRetr, GridcolCardGb, "CD", "NM", "");
            RepositoryItemGridLookUpEdit cardKindup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(cardKindup, dtCardKind, GridRetr, GridColKind, "CD", "NM", "");
            RepositoryItemGridLookUpEdit cardCoCdLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(cardCoCdLkup, dtCardCompany, GridRetr, GridColCoCd, "CD", "NM", "");
            
            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            
            BtnRetr_Click(null, null);
        }
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("3"))
            {
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'CARD_GB'");

            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'CARD_CO_CD'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'CARD_KIND'");
            }
            else
            {
                strSql.AppendLine(" SELECT A.ACC_CD AS CD");
                strSql.AppendLine("      , A.ACC_NM AS NM");
                strSql.AppendLine("   FROM ACC_ACC_CD A");
            }
            if (sOther.Equals("Y"))
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

            GetGridRetr();
        }
        private void GetGridRetr()
        {
            Cursor = Cursors.WaitCursor;

            GridRetr.DataSource = null;

            StringBuilder strSql = new StringBuilder();

            string sCardGB = LkupEditCardGB.EditValue.ToString();
            string sCardCompany = LkupEditCardCompany.EditValue.ToString();
            string sUseYn = RdGrUseYn.EditValue.ToString();
            string sCardNo = TxtCardNo.Text.Replace("_", "").Replace("-", "");

            strSql.Clear();
            strSql.AppendLine("");
            strSql.AppendLine(" SELECT A.CARD_GB  ");//카드구분
            strSql.AppendLine(" 	 , A.CARD_NO ");//카드번호
            strSql.AppendLine(" 	 , A.CARD_NM");//카드명
            strSql.AppendLine(" 	 , A.CARD_KIND");//카드종류
            strSql.AppendLine(" 	 , A.CARD_CO_CD");//카드회사코드
            strSql.AppendLine(" 	 , A.OWNER_DEPT");//사용부서
            strSql.AppendLine(" 	 , A.OWNER_ID");//카드사용자
            strSql.AppendLine(" 	 , A.PRJT_NO");//연구번호
            strSql.AppendLine(" 	 , A.LMT_AMT");//한도금액
            strSql.AppendLine(" 	 , A.DISUSE_YMD");//폐기일자
            strSql.AppendLine(" 	 , A.USE_YN");// 사용여부
            strSql.AppendLine(" 	 , A.NOTE");//비고         
            strSql.AppendLine(" 	 , A.ENT_DT");//입력일자
            strSql.AppendLine(" 	 , A.ENT_ID");//입력ID
            strSql.AppendLine(" 	 , A.MFY_DT");//수정일자
            strSql.AppendLine(" 	 , A.MFY_ID");//수정ID
            strSql.AppendLine("	  FROM ACC_CARD_CD A");
            strSql.AppendLine("	 WHERE 1=1");
            strSql.AppendLine("    AND (('' = '" + sCardGB + "') OR (('' <> '" + sCardGB + "') AND A.CARD_GB = '" + sCardGB + "'))");
            strSql.AppendLine("    AND (('' = '" + sCardCompany + "') OR (('' <> '" + sCardCompany + "') AND A.CARD_CO_CD = '" + sCardCompany + "'))");
            strSql.AppendLine("    AND (('' = '" + sUseYn + "') OR (('' <> '" + sUseYn + "') AND A.USE_YN = '" + sUseYn + "'))");
            if (sCardNo != "")
            {
                strSql.AppendLine("AND CARD_NO = '" + sCardNo + "'");
            }


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridRetr.DataSource = dt;
            Cursor = Cursors.Default;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if(GridViewRetr.RowCount > 0)
            {
                string sGb = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, "CARD_GB")?.ToString();
                string sNo = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, "CARD_NO")?.ToString();

                if (string.IsNullOrEmpty(sGb))
                {
                    XtraMessageBox.Show("카드구분을 입력하세요.");
                    return;
                }
                else if (string.IsNullOrEmpty(sNo))
                {
                    XtraMessageBox.Show("카드번호를 입력하세요.");
                    return;
                }
            }

            GridViewRetr.AddNewRow();
            GridViewRetr.Focus();
            GridViewRetr.FocusedColumn = GridcolCardGb;

            GridViewRetr.SetFocusedRowCellValue("USE_YN", "Y");
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)GridRetr.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sGb = dtMerge.Rows[i]["CARD_GB"]?.ToString();
                string sNo = dtMerge.Rows[i]["CARD_NO"]?.ToString();

                if (string.IsNullOrEmpty(sGb) || string.IsNullOrEmpty(sNo))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            string sCardGb = string.Empty;
            string sCardNo = string.Empty;
            string sCardNm = string.Empty;
            string sCardKind = string.Empty;
            string sCardCoCd = string.Empty;
            string sOwnerDept = string.Empty;
            string sOwnerId = string.Empty;
            string sPrjtNo = string.Empty;
            double sLmtAmt = 0;
            string sDisuseYmd = string.Empty;
            string sUseYn = string.Empty;
            string sNote = string.Empty;
            string sEntDt = string.Empty;
            string sId = FmMainToolBar2.drUser["USRCD"]?.ToString();

            if (dtMerge.Rows.Count > 0)
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    StringBuilder strSql = new StringBuilder();

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        if (dtMerge.Rows[j]["CARD_GB"] is DBNull || dtMerge.Rows[j]["CARD_NO"] is DBNull || dtMerge.Rows[j]["CARD_CO_CD"] is DBNull)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            Cursor = Cursors.Default;
                            MessageBox.Show("카드구분,카드번호,카드회사를 입력해주세요");
                            return;
                        }
                        sCardGb = Convert.ToString(dtMerge.Rows[j]["CARD_GB"]);
                        sCardNo = Convert.ToString(dtMerge.Rows[j]["CARD_NO"]).Replace("-", "");

                        sCardNm = dtMerge.Rows[j]["CARD_NM"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["CARD_NM"]);
                        sCardKind = dtMerge.Rows[j]["CARD_KIND"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["CARD_KIND"]);
                        sCardCoCd = dtMerge.Rows[j]["CARD_CO_CD"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["CARD_CO_CD"]);
                        sOwnerDept = dtMerge.Rows[j]["OWNER_DEPT"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["OWNER_DEPT"]);
                        sOwnerId = dtMerge.Rows[j]["OWNER_ID"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["OWNER_ID"]);
                        sPrjtNo = dtMerge.Rows[j]["PRJT_NO"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["PRJT_NO"]);
                        sLmtAmt = dtMerge.Rows[j]["LMT_AMT"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["LMT_AMT"]);
                        if (dtMerge.Rows[j]["DISUSE_YMD"].ToString().Length > 1)
                        {
                            sDisuseYmd = dtMerge.Rows[j]["DISUSE_YMD"] is DBNull ? "" : Convert.ToString(dtMerge.Rows[j]["DISUSE_YMD"]).Replace("-", "").Substring(0, 8);
                        }
                        else { sDisuseYmd = ""; }                   
                        
                        sUseYn = dtMerge.Rows[j]["USE_YN"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["USE_YN"]);
                        sNote = dtMerge.Rows[j]["NOTE"] is null ? "" : Convert.ToString(dtMerge.Rows[j]["NOTE"]);



                        strSql.Clear();
                        strSql.AppendLine("");

                        strSql.AppendLine("  MERGE INTO ACC_CARD_CD AS a                 ");
                        strSql.AppendLine("      USING(SELECT CARD_GB = '" + sCardGb + "'               ");
                        strSql.AppendLine("             , CARD_NO = '" + sCardNo + "'                   ");
                        strSql.AppendLine("             , CARD_NM = '" + sCardNm + "'                   ");
                        strSql.AppendLine("             , CARD_KIND = '" + sCardKind + "'               ");
                        strSql.AppendLine("             , CARD_CO_CD = '" + sCardCoCd + "'             ");
                        strSql.AppendLine("             , OWNER_DEPT = '" + sOwnerDept + "'                ");
                        strSql.AppendLine("             , OWNER_ID = '" + sOwnerId + "'            ");
                        strSql.AppendLine("             , PRJT_NO = '" + sPrjtNo + "'                   ");
                        strSql.AppendLine("             , LMT_AMT = '" + sLmtAmt + "'                  ");
                        strSql.AppendLine("               , DISUSE_YMD = '" + sDisuseYmd + "'      ");
                        strSql.AppendLine("             , USE_YN = '" + sUseYn + "'                   ");
                        strSql.AppendLine("             , NOTE = '" + sNote + "'                      ");
                        strSql.AppendLine("             , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20)             ");
                        strSql.AppendLine("             , MFY_ID = '"+ sId + "') AS b        ");
                        strSql.AppendLine("       ON (a.CARD_GB = b.CARD_GB and a.CARD_NO = b.CARD_NO)               ");
                        strSql.AppendLine("       WHEN MATCHED THEN UPDATE SET           ");
                        strSql.AppendLine("               CARD_GB = '" + sCardGb + "'                   ");
                        strSql.AppendLine("             , CARD_NO = '" + sCardNo + "'                   ");
                        strSql.AppendLine("             , CARD_NM = '" + sCardNm + "'                   ");
                        strSql.AppendLine("             , CARD_KIND = '" + sCardKind + "'               ");
                        strSql.AppendLine("             , CARD_CO_CD = '" + sCardCoCd + "'             ");
                        strSql.AppendLine("             , OWNER_DEPT = '" + sOwnerDept + "'                ");
                        strSql.AppendLine("             , OWNER_ID = '" + sOwnerId + "'            ");
                        strSql.AppendLine("             , PRJT_NO = '" + sPrjtNo + "'                   ");
                        strSql.AppendLine("             , LMT_AMT = '" + sLmtAmt + "'                  ");
                        strSql.AppendLine("               , DISUSE_YMD = '" + sDisuseYmd + "'      ");
                        strSql.AppendLine("             , USE_YN = '" + sUseYn + "'                   ");
                        strSql.AppendLine("             , NOTE = '" + sNote + "'                      ");
                        strSql.AppendLine("             , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20)              ");
                        strSql.AppendLine("             , MFY_ID = '"+sId+"'              ");
                        strSql.AppendLine("       WHEN NOT MATCHED THEN INSERT(          ");
                        strSql.AppendLine("               CARD_GB                        ");
                        strSql.AppendLine("             , CARD_NO                        ");
                        strSql.AppendLine("             , CARD_NM                        ");
                        strSql.AppendLine("             , CARD_KIND                      ");
                        strSql.AppendLine("             , CARD_CO_CD                     ");
                        strSql.AppendLine("             , OWNER_DEPT                     ");
                        strSql.AppendLine("             , OWNER_ID                       ");
                        strSql.AppendLine("             , PRJT_NO                        ");
                        strSql.AppendLine("             , LMT_AMT                        ");
                        strSql.AppendLine("             , DISUSE_YMD                     ");
                        strSql.AppendLine("             , USE_YN                         ");
                        strSql.AppendLine("             , NOTE                           ");
                        strSql.AppendLine("             , ENT_DT                         ");
                        strSql.AppendLine("             , ENT_ID)                        ");
                        strSql.AppendLine("          VALUES('" + sCardGb + "'                           ");
                        strSql.AppendLine("             , '" + sCardNo + "'                             ");
                        strSql.AppendLine("             , '" + sCardNm + "'                             ");
                        strSql.AppendLine("             , '" + sCardKind + "'                           ");
                        strSql.AppendLine("             , '" + sCardCoCd + "'                          ");
                        strSql.AppendLine("             , '" + sOwnerDept + "'                             ");
                        strSql.AppendLine("             , '" + sOwnerId + "'                       ");
                        strSql.AppendLine("             , '" + sPrjtNo + "'                             ");
                        strSql.AppendLine("             , '" + sLmtAmt + "'                            ");
                        strSql.AppendLine("             , '" + sDisuseYmd + "'                     ");
                        strSql.AppendLine("             , '" + sUseYn + "'                            ");
                        strSql.AppendLine("             , '" + sNote + "'                             ");
                        strSql.AppendLine("             , CONVERT(VARCHAR(20),GETDATE(),20)                       ");
                        strSql.AppendLine("             , '"+sId+"'                       ");
                        strSql.AppendLine(" );"      );





                    /*    
                        strSql.AppendLine("INSERT INTO ACC_CARD_CD ");
                        strSql.AppendLine("           (CARD_GB       ");
                        strSql.AppendLine("           ,CARD_NO ");
                        strSql.AppendLine("           ,CARD_NM ");
                        strSql.AppendLine("           ,CARD_KIND ");
                        strSql.AppendLine("           ,CARD_CO_CD      ");
                        strSql.AppendLine("           ,OWNER_DEPT       ");
                        strSql.AppendLine("           ,OWNER_ID    ");
                        strSql.AppendLine("           ,PRJT_NO      ");
                        strSql.AppendLine("           ,LMT_AMT    ");
                        strSql.AppendLine(" 	      ,DISUSE_YMD     ");
                        strSql.AppendLine(" 	      ,USE_YN       ");
                        strSql.AppendLine(" 	      ,NOTE       ");
                        strSql.AppendLine(" 	      ,ENT_DT      ");
                        strSql.AppendLine(" 	      ,ENT_ID       ");
                        strSql.AppendLine(" 	      ,MFY_DT       ");
                        strSql.AppendLine(" 	      ,MFY_ID )     ");
                        strSql.AppendLine("     VALUES (");
                        strSql.AppendLine("            '" + sCardGb + "'");
                        strSql.AppendLine(" 		  ,'" + sCardNo + "'   ");
                        strSql.AppendLine(" 		  ,'" + sCardNm + "'   ");
                        strSql.AppendLine(" 		  ,'" + sCardKind + "'   ");
                        strSql.AppendLine(" 		  ,'" + sCardCoCd + "'    ");
                        strSql.AppendLine(" 		  ,'" + sOwnerDept + "'     ");
                        strSql.AppendLine(" 		  ,'" + sOwnerId + "'  ");
                        strSql.AppendLine(" 		  ,'" + sPrjtNo + "'    ");
                        strSql.AppendLine(" 		  ,'" + sLmtAmt + "'  ");
                        strSql.AppendLine(" 		  ,'" + sDisuseYmd + "'   ");
                        strSql.AppendLine(" 		  ,'" + sUseYn + "'    ");
                        strSql.AppendLine(" 		  ,'" + sNote + "'    ");
                        strSql.AppendLine(" 		  ,NOW()    ");
                        strSql.AppendLine(" 		  , '입력ID'   ");
                        strSql.AppendLine(" 		  ,NOW()     ");
                        strSql.AppendLine(" 		  , '입력ID동일'  ");
                        strSql.AppendLine("            )                    ");
                        strSql.AppendLine("         ON DUPLICATE KEY UPDATE              ");
                        strSql.AppendLine("  		   CARD_NM             ='" + sCardNm + "'     ");
                        strSql.AppendLine("           ,CARD_KIND           ='" + sCardKind + "'  ");
                        strSql.AppendLine(" 	      ,CARD_CO_CD          ='" + sCardCoCd + "'    ");
                        strSql.AppendLine(" 	      ,OWNER_DEPT          ='" + sOwnerDept + "'  ");
                        strSql.AppendLine(" 	      ,OWNER_ID            ='" + sOwnerId + "'   ");
                        strSql.AppendLine(" 	      ,PRJT_NO             ='" + sPrjtNo + "'    ");
                        strSql.AppendLine(" 	      ,LMT_AMT             ='" + sLmtAmt + "'     ");
                        strSql.AppendLine("   	      ,DISUSE_YMD          ='" + sDisuseYmd + "'    ");
                        strSql.AppendLine(" 	      ,USE_YN              ='" + sUseYn + "'    ");
                        strSql.AppendLine(" 	      ,NOTE                ='" + sNote + "'     ");                     
                        strSql.AppendLine(" 	      ,MFY_DT              = NOW()    ");
                        strSql.AppendLine(" 	      ,MFY_ID              ='수정ID'    ");

                        */

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "Table:ACC_CARD_CD -> CARD_GB:" + sCardGb + ",CARD_NO:" + sCardNo + ",CARD_NM:" + sCardNm;
                        ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, (j + 1).ToString(), "S", this.Name, sLogRmk, cmd);
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridRetr();
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColCardNo, dtMerge.Rows[0]["CARD_NO"]?.ToString());
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                    Cursor = Cursors.Default;
                }
            }
            else
            {
                XtraMessageBox.Show("변경된 내용이 없습니다.");
            }

            Cursor = Cursors.Default;
        }

        private void repoDateEdit_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value.ToString() == "") return;
            string sYmd = e.Value.ToString().Replace("-", "").Substring(0, 8);
            e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
        }

        private void GridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.Equals("GridColDisuseYmd"))
            {
                string sYmd = e.Value.ToString().Replace("-", "").Substring(0, 8);
                e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewRetr_ShowingPopupEditForm(object sender, DevExpress.XtraGrid.Views.Grid.ShowingPopupEditFormEventArgs e)
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

        private void AccCreditCardMgt_KeyDown(object sender, KeyEventArgs e)
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

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void AccCreditCardMgt_FormClosed(object sender, FormClosedEventArgs e)
        {
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "AccCreditCardMgt", GridViewRetr);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sCardGb = GridViewRetr.GetFocusedRowCellValue("CARD_GB")?.ToString();
            string sCardNo = GridViewRetr.GetFocusedRowCellValue("CARD_NO")?.ToString();

            if (XtraMessageBox.Show("카드구분 : " + sCardGb + "\r\n카드번호 : " + sCardNo
                + "\r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
               , "신용카드 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            if(string.IsNullOrEmpty(sCardGb) || string.IsNullOrEmpty(sCardNo))
            {
                XtraMessageBox.Show("삭제 필수사항이 존재하지 않습니다.");
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
                strSql.AppendLine(" DELETE FROM ACC_CARD_CD ");
                strSql.AppendLine("       WHERE CARD_GB = '" + sCardGb + "' ");
                strSql.AppendLine("         AND CARD_NO = '" + sCardNo + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:ACC_CARD_CD -> CARD_GB:" + sCardGb + ",CARD_NO:" + sCardNo;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;

                BtnRetr.PerformClick();

                GridViewRetr.FocusedRowHandle = idx;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
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
                string sFileNM = "신용카드 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridRetr.ExportToXls(FileName + ".xls");
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

        private void AccCreditCardMgt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr.UpdateCurrentRow();
        }

        private void LkupEditCardCompany_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}