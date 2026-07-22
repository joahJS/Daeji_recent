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
using System.Diagnostics;
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
    public partial class CommonCdMgt : DevExpress.XtraEditors.XtraForm
    {
        public CommonCdMgt()
        {
            InitializeComponent();
        }
      
        private string GetGb;
        private string GetNm;

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void CommonCdMgt_Load(object sender, EventArgs e)
        {
            //
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "CommonCdMgt");
            DataTable dtComCd = GetLookUpData("1", "1", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupCodeName, dtComCd, "CD", "NM", "Y");
            DataTable dtUserCd = GetLookUpData("2", "", "", "");
            RepositoryItemGridLookUpEdit userLkup = new RepositoryItemGridLookUpEdit(); 
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GrideCode, GridColTopMfyId, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GridCodeDept, GridColMfyId, "CD", "NM", "");
            // SetCombo();
            arrGrdView = new GridView[] { GridViewCode, GridViewCodeDept };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            BtnRetr.PerformClick();
        }
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine(" UNION ALL");
            }


            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'COM_COM_CD'");

            }
            if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.USRCD AS CD");
                strSql.AppendLine("      , A.USRNM AS NM");
                strSql.AppendLine("      , A.USRCD AS SEQ");
                strSql.AppendLine("   FROM ZUSRLST A");
                strSql.AppendLine("  WHERE USEYN = 'Y'");

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

        private void SetCombo()
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //StringBuilder strSql = new StringBuilder();
            //strSql.Clear();
            //strSql.AppendLine("");
            //strSql.AppendLine("SELECT COM_CD as Items");
            //strSql.AppendLine("      ,COM_NM    ");
            //strSql.AppendLine(" FROM COM_BASE_CD");
            //strSql.AppendLine("WHERE CD_GB='COM_COM_CD' ");
            //strSql.AppendLine("GROUP BY Items");
            //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    string Item = dt.Rows[i]["COM_NM"].ToString()+"  :  "+dt.Rows[i]["Items"].ToString();
            //    CboCodeName.Properties.Items.Add(Item);
            //}
        }
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            GetGridTop();
            //GridCodeDept.DataSource = null;
        }
        private void GetGridTop()
        {
            GrideCode.DataSource = null;
            StringBuilder strSql = new StringBuilder();

            string sFindWord = TxtFindWord.EditValue?.ToString();

            strSql.AppendLine("");
            strSql.AppendLine(" SELECT A.CD_GB  ");//visible false
            strSql.AppendLine(" 	 , A.CD_GB_NM ");//visible false
            strSql.AppendLine(" 	 , A.COM_CD ");
            strSql.AppendLine(" 	 , A.COM_NM ");
            strSql.AppendLine(" 	 , A.USE_YN ");
            strSql.AppendLine(" 	 , A.SORT_SEQ ");
            strSql.AppendLine(" 	 , A.REMARK ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine("	 FROM COM_BASE_CD A");
            strSql.AppendLine(" WHERE 1=1");
            strSql.AppendLine("   AND CD_GB ='COM_COM_CD'");
            strSql.AppendLine("   AND COM_NM LIKE '%" + sFindWord + "%'");
            strSql.AppendLine(" ORDER BY SORT_SEQ ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count < 1)
            {
                MessageBox.Show("조회결과가 없습니다");
            }
            else
            {
                GrideCode.DataSource = dt;
            }

            //자동순번 초기화
            iSeqCnt1 = 1;
            iSeqCnt2 = 1;
        }

        int iSeqCnt1 = 1;
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sComcd = GridViewCode.GetRowCellValue(GridViewCode.RowCount - 1, "COM_CD")?.ToString();

            if (string.IsNullOrEmpty(sComcd))
                return;

            GridViewCode.Focus();
            GridViewCode.AddNewRow();
            GridViewCode.SetFocusedRowCellValue("USE_YN", "Y");
            GridViewCode.FocusedColumn = GridColTopComCd;

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT MAX(SORT_SEQ) AS SORT_SEQ");
            strSql.AppendLine("	  FROM COM_BASE_CD A");
            strSql.AppendLine("  WHERE CD_GB ='COM_COM_CD'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt != null && dt.Rows.Count > 0)
            {
                string sSORT_SEQ = dt.Rows[0]["SORT_SEQ"]?.ToString();

                if(int.TryParse(sSORT_SEQ, out int iResult))
                {
                    GridViewCode.SetFocusedRowCellValue("SORT_SEQ", iResult + iSeqCnt1++);
                }
                else
                {
                    GridViewCode.SetFocusedRowCellValue("SORT_SEQ", iSeqCnt1++);
                }
            }

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            SaveGridTop();
        }
        private void SaveGridTop()
        {
            //if(GridViewCodeDept.RowCount == 0)
            //{
            //    XtraMessageBox.Show("상세내역에 코드를 하나이상 입력하세요.");
            //    return;
            //}

            DataTable dt = (DataTable)GrideCode.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sCOM_CD = dtMerge.Rows[i]["COM_CD"]?.ToString();

                if (string.IsNullOrEmpty(sCOM_CD))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            string sCdGb = string.Empty;
            string sCdGbNm = string.Empty;
            string sComCd = string.Empty;
            string sComNm = string.Empty;
            string sUseYn = string.Empty;
            double sSortSeq = 0;
            string sRemark = string.Empty;
            string sMfyDt = string.Empty;
            string sMfyId = string.Empty;
            
            StringBuilder strSql = new StringBuilder();
            try
            {
                if (dtMerge.Rows.Count > 0)
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        if (dtMerge.Rows[j]["COM_CD"]is DBNull)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("공통코드 를 입력하세요");
                            return;
                        }
                        if (dtMerge.Rows[j]["COM_NM"] is DBNull)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("공통코드명 을 입력하세요");
                            return;
                        }
                        if (dtMerge.Rows[j]["SORT_SEQ"] is DBNull)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("순번 을 입력하세요");
                            return;
                        }

                        sCdGb = "COM_COM_CD";
                        sCdGbNm = "공통코드 구분";
                        sComCd = Convert.ToString(dtMerge.Rows[j]["COM_CD"]);
                        sComNm = Convert.ToString(dtMerge.Rows[j]["COM_NM"]);
                        sUseYn = Convert.ToString(dtMerge.Rows[j]["USE_YN"]);
                        sSortSeq = Convert.ToDouble(dtMerge.Rows[j]["SORT_SEQ"]);
                        sRemark = Convert.ToString(dtMerge.Rows[j]["REMARK"]);
                        string sUSERCD = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine("IF EXISTS(SELECT* FROM COM_BASE_CD WHERE CD_GB = '"+ sCdGb + "' AND COM_CD = '"+ sComCd + "') ");
                        strSql.AppendLine("    BEGIN                                                            ");
                        strSql.AppendLine("          UPDATE COM_BASE_CD                                         ");
                        strSql.AppendLine("             SET COM_NM = '" + sComNm + "'                           ");
                        strSql.AppendLine("               , USE_YN = '" + sUseYn + "'                           ");
                        strSql.AppendLine("               , SORT_SEQ = '" + sSortSeq + "'                       ");
                        strSql.AppendLine("               , REMARK = '" + sRemark + "'                          ");
                        strSql.AppendLine("               , MFY_ID = '" + sUSERCD + "'                          ");
                        strSql.AppendLine("               , MFY_DT = CONVERT(DATE, GETDATE())                                  ");
                        strSql.AppendLine("           WHERE CD_GB = '"+ sCdGb + "' AND COM_CD = '"+ sComCd + "' ");
                        strSql.AppendLine("      END                                                            ");
                        strSql.AppendLine(" ELSE                                                                ");
                        strSql.AppendLine("    BEGIN                                                            ");
                        strSql.AppendLine("           INSERT INTO COM_BASE_CD                                   ");
                        strSql.AppendLine("                ( CD_GB                                               ");
                        strSql.AppendLine("                , CD_GB_NM                                           ");
                        strSql.AppendLine("                , COM_CD                                             ");
                        strSql.AppendLine("                , COM_NM                                             ");
                        strSql.AppendLine("                , USE_YN                                            ");
                        strSql.AppendLine("                , SORT_SEQ                                          ");
                        strSql.AppendLine("                , REMARK                                            ");
                        strSql.AppendLine("                , MFY_ID                                            ");
                        strSql.AppendLine("                , MFY_DT)                                           ");
                        strSql.AppendLine("         VALUES(                                                     ");
                        strSql.AppendLine("                '" + sCdGb + "'                                      ");
                        strSql.AppendLine("                    , '" + sCdGbNm + "'                              ");
                        strSql.AppendLine("                    , '" + sComCd + "'                               ");
                        strSql.AppendLine("                    , '" + sComNm + "'                               ");
                        strSql.AppendLine("                    , '" + sUseYn + "'                               ");
                        strSql.AppendLine("                    , '" + sSortSeq + "'                             ");
                        strSql.AppendLine("                    , '" + sRemark + "'                              ");
                        strSql.AppendLine("                    , '" + sUSERCD + "'                              ");
                        strSql.AppendLine("                    , CONVERT(DATE, GETDATE())                                     ");
                        strSql.AppendLine("                )                                                    ");
                        strSql.AppendLine("      END                                                            ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "CD_GB:" + sCdGb + ",CD_GB_NM:" + sCdGbNm + ",COM_CD:" + sComCd + ",COM_NM:" + sComNm;
                        //ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, (j+1).ToString(), "S", this.Name, sLogRmk, cmd);

                        string sIP = ClsFunc.GetLocalIP();

                        strSql.Clear();
                        strSql.AppendLine(" INSERT INTO  ZSYS_LOG ");
                        strSql.AppendLine(" 	         ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                        strSql.AppendLine(" 	    VALUES ");
                        strSql.AppendLine(" 	         ( '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ");
                        strSql.AppendLine(" 	         , '" + FmMainToolBar2.UserID + "'  ");
                        strSql.AppendLine(" 	         , " + (j + 1).ToString() + " ");
                        strSql.AppendLine(" 	         , '" + this.Name + "' ");
                        strSql.AppendLine(" 	         , 'S' ");
                        strSql.AppendLine(" 	         , '" + sIP + "' ");
                        strSql.AppendLine(" 	         , '" + sLogRmk + "' ) ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridTop();
                    GridViewCode.FocusedRowHandle = GridViewCode.LocateByDisplayText(0, GridColComCd, sComCd);
                }
                else
                {
                    XtraMessageBox.Show("변경된 데이터가 없습니다.");
                }

            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }
        private void CommonCdMgt_KeyDown(object sender, KeyEventArgs e)
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
            else if (e.KeyCode == Keys.F9)
            {
                BtnDeptAdd.PerformClick();
            }
            else if (e.KeyCode == Keys.F11)
            {
                BtnDeptSave.PerformClick();
            }
            else if (e.KeyCode == Keys.F12)
            {
                BtnDeptDelete.PerformClick();
            }
        }

        private void GrideCode_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridCodeDept_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewCode_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            
        }
        private void GetGridCodeDept(string sGb)
        {

            GridCodeDept.DataSource = null;
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("");
            strSql.AppendLine(" SELECT A.CD_GB  ");//visible false
            strSql.AppendLine(" 	 , A.CD_GB_NM ");//visible false
            strSql.AppendLine(" 	 , A.COM_CD ");
            strSql.AppendLine(" 	 , A.COM_NM ");
            strSql.AppendLine(" 	 , A.REMARK ");
            strSql.AppendLine(" 	 , A.COM_SUB_CD1 ");
            strSql.AppendLine(" 	 , A.COM_SUB_NM1 ");
            strSql.AppendLine(" 	 , A.COM_SUB_CD2 ");
            strSql.AppendLine(" 	 , A.COM_SUB_NM2 ");
            strSql.AppendLine(" 	 , A.COM_SUB_CD3 ");
            strSql.AppendLine(" 	 , A.COM_SUB_NM3 ");
            strSql.AppendLine(" 	 , A.CD_AMT1 ");
            strSql.AppendLine(" 	 , A.CD_AMT2 ");
            strSql.AppendLine(" 	 , A.CD_AMT3 ");
            strSql.AppendLine(" 	 , A.CD_YN1 ");
            strSql.AppendLine(" 	 , A.CD_YN2 ");
            strSql.AppendLine(" 	 , A.CD_YN3 ");
            strSql.AppendLine(" 	 , A.SORT_SEQ ");
            strSql.AppendLine(" 	 , A.USE_YN ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("	 FROM COM_BASE_CD A");
            strSql.AppendLine(" WHERE A.CD_GB = '" + sGb + "'");
            strSql.AppendLine(" ORDER BY A.SORT_SEQ ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridCodeDept.DataSource = dt;

            //자동순번 초기화
            iSeqCnt2 = 1;
        }

        int iSeqCnt2 = 1;
        private void BtnDeptAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sComcd = GridViewCodeDept.GetRowCellValue(GridViewCodeDept.RowCount - 1, "COM_CD")?.ToString();

            if (GridViewCodeDept.RowCount > 0 && string.IsNullOrEmpty(sComcd))
                return;

            GridViewCodeDept.AddNewRow();
            GridViewCodeDept.SetFocusedRowCellValue("USE_YN", "Y");
            GridViewCodeDept.FocusedColumn = GridColComCd;

            string sCOMCD = GridViewCode.GetFocusedRowCellValue("COM_CD")?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT MAX(SORT_SEQ) AS SORT_SEQ");
            strSql.AppendLine("	 FROM COM_BASE_CD A");
            strSql.AppendLine(" WHERE CD_GB ='"+ sCOMCD + "'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt != null && dt.Rows.Count > 0)
            {
                string sSORT_SEQ = dt.Rows[0]["SORT_SEQ"]?.ToString();

                if (int.TryParse(sSORT_SEQ, out int iResult))
                {
                    GridViewCodeDept.SetFocusedRowCellValue("SORT_SEQ", iResult+ iSeqCnt2++);
                }
                else
                {
                    GridViewCodeDept.SetFocusedRowCellValue("SORT_SEQ", iSeqCnt2++);
                }
            }
        }

        private void BtnDeptDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            // if (GridViewRetr.GetFocusedRowCellValue("J_Serial") == null) return;
            if ((MessageBox.Show(this, "삭제 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
            {
                return;
            }
            DeleteCodeDept();
           
        }
        private void DeleteCodeDept()
        {
            int idx = GridViewCodeDept.FocusedRowHandle;

            string sComGb = string.Empty;
            sComGb = GridViewCodeDept.GetFocusedRowCellValue("CD_GB").ToString();

            string sComCd = string.Empty;
            sComCd = GridViewCodeDept.GetFocusedRowCellValue("COM_CD").ToString();
            StringBuilder strSql = new StringBuilder();
            
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine("DELETE FROM com_base_cd ");
                strSql.AppendLine(" WHERE CD_GB='" + sComGb + "' ");
                strSql.AppendLine("   AND COM_CD='" + sComCd + "'");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "CD_GB:" + sComGb + ",COM_CD:" + sComCd;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                GetGridCodeDept(sComGb);
                GridViewCodeDept.FocusedRowHandle = idx-1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnDeptSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sGb = GetGb;
            string sGbNm = GetNm;
            SaveCodeDept(sGb, sGbNm);
        }

        private void SaveCodeDept(string sGb, string sGbNm)
        {
            DataTable dt = (DataTable)GridCodeDept.DataSource;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sCOM_CD = dtMerge.Rows[i]["COM_CD"]?.ToString();

                if (string.IsNullOrEmpty(sCOM_CD))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            string sCdGb = string.Empty;
            string sComCd = string.Empty;
            string sComSubCd1 = string.Empty;
            string sComSubCd2 = string.Empty;
            string sComSubCd3 = string.Empty;
            string sCdGbNm = string.Empty;
            string sComNm = string.Empty;
            string sComSubNm1 = string.Empty;
            string sComSubNm2 = string.Empty;
            string sComSubNm3 = string.Empty;
            double sCdAmt1 = 0;
            double sCdAmt2 = 0;
            double sCdAmt3 = 0;
            string sCdYn1 = string.Empty;
            string sCdYn2 = string.Empty;
            string sCdYn3 = string.Empty;
            double sSortSeq = 0;
            string sUseYn = string.Empty;
            string sMfyDt = string.Empty;
            string sMfyId = string.Empty;
            string sRemark = string.Empty;
            
            if(dtMerge.Rows.Count == 0)
            {
                MessageBox.Show("저장을 완료했습니다.");
                return;
            }
         
            try
            {
                if (dtMerge.Rows.Count > 0)
                {
                    StringBuilder strSql = new StringBuilder();

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        if (dtMerge.Rows[j]["COM_CD"] is DBNull)
                        {
                            MessageBox.Show("코드란 을 입력하세요");
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }
                        if (dtMerge.Rows[j]["COM_NM"] is DBNull)
                        {
                            MessageBox.Show("코드명란 을 입력하세요");
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }
                        if (dtMerge.Rows[j]["SORT_SEQ"] is DBNull)
                        {
                            MessageBox.Show("순번 을 입력하세요");
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            return;
                        }

                        sCdGb = sGb;
                        sComCd = Convert.ToString(dtMerge.Rows[j]["COM_CD"]);
                        sComSubCd1 = Convert.ToString(dtMerge.Rows[j]["COM_SUB_CD1"]);
                        sComSubCd2 = Convert.ToString(dtMerge.Rows[j]["COM_SUB_CD2"]);
                        sComSubCd3 = Convert.ToString(dtMerge.Rows[j]["COM_SUB_CD3"]);
                        sCdGbNm = sGbNm;
                        sComNm = Convert.ToString(dtMerge.Rows[j]["COM_NM"]);
                        sComSubNm1 = Convert.ToString(dtMerge.Rows[j]["COM_SUB_NM1"]);
                        sComSubNm2 = Convert.ToString(dtMerge.Rows[j]["COM_SUB_NM2"]);
                        sComSubNm3 = Convert.ToString(dtMerge.Rows[j]["COM_SUB_NM3"]);
                        sCdAmt1 = dtMerge.Rows[j]["CD_AMT1"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["CD_AMT1"]);
                        sCdAmt2 = dtMerge.Rows[j]["CD_AMT2"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["CD_AMT2"]);
                        sCdAmt3 = dtMerge.Rows[j]["CD_AMT3"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["CD_AMT3"]);
                        sCdYn1 = Convert.ToString(dtMerge.Rows[j]["CD_YN1"]);
                        sCdYn2 = Convert.ToString(dtMerge.Rows[j]["CD_YN2"]);
                        sCdYn3 = Convert.ToString(dtMerge.Rows[j]["CD_YN3"]);
                        sSortSeq = dtMerge.Rows[j]["SORT_SEQ"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SORT_SEQ"]);
                        sUseYn = Convert.ToString(dtMerge.Rows[j]["USE_YN"]);
                        sMfyDt = Convert.ToString(dtMerge.Rows[j]["MFY_DT"]);
                        sMfyId = Convert.ToString(dtMerge.Rows[j]["MFY_ID"]);
                        sRemark = Convert.ToString(dtMerge.Rows[j]["REMARK"]);

                        string sUSERCD = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine("");
                        strSql.AppendLine("IF EXISTS(SELECT* FROM COM_BASE_CD WHERE CD_GB = '" + sCdGb + "' AND COM_CD = '" + sComCd + "')");
                        strSql.AppendLine("    BEGIN                                                                                      ");
                        strSql.AppendLine("          UPDATE COM_BASE_CD                                                                   ");
                        strSql.AppendLine("             SET COM_SUB_CD1 = '" + sComSubCd1 + "'                                            ");
                        strSql.AppendLine("               , COM_SUB_CD2 = '" + sComSubCd2 + "'                                            ");
                        strSql.AppendLine("               , COM_SUB_CD3 = '" + sComSubCd3 + "'                                            ");
                        strSql.AppendLine("               , CD_GB_NM = '" + sCdGbNm + "'                                                  ");
                        strSql.AppendLine("               , COM_NM = '" + sComNm + "'                                                     ");
                        strSql.AppendLine("               , COM_SUB_NM1 = '" + sComSubNm1 + "'                                            ");
                        strSql.AppendLine("               , COM_SUB_NM2 = '" + sComSubNm2 + "'                                            ");
                        strSql.AppendLine("               , COM_SUB_NM3 = '" + sComSubNm3 + "'                                            ");
                        strSql.AppendLine("               , CD_AMT1 = '" + sCdAmt1 + "'                                                   ");
                        strSql.AppendLine("               , CD_AMT2 = '" + sCdAmt2 + "'                                                   ");
                        strSql.AppendLine("               , CD_AMT3 = '" + sCdAmt3 + "'                                                   ");
                        strSql.AppendLine("               , CD_YN1 = '" + sCdYn1 + "'                                                     ");
                        strSql.AppendLine("               , CD_YN2 = '" + sCdYn2 + "'                                                     ");
                        strSql.AppendLine("               , CD_YN3 = '" + sCdYn3 + "'                                                     ");
                        strSql.AppendLine("               , SORT_SEQ = '" + sSortSeq + "'                                                 ");
                        strSql.AppendLine("               , USE_YN = '" + sUseYn + "'                                                     ");
                        strSql.AppendLine("               , MFY_DT = CONVERT(DATE, GETDATE())                                                           ");
                        strSql.AppendLine("               , MFY_ID = '" + sUSERCD + "'                                                    ");
                        strSql.AppendLine("               , REMARK = '" + sRemark + "'                                                    ");
                        strSql.AppendLine("           WHERE CD_GB = '" + sCdGb + "' AND COM_CD = '" + sComCd + "'                                                    ");
                        strSql.AppendLine("      END                                                                                      ");
                        strSql.AppendLine(" ELSE                                                                                          ");
                        strSql.AppendLine("    BEGIN                                                                                      ");
                        strSql.AppendLine("           INSERT INTO COM_BASE_CD                                                             ");
                        strSql.AppendLine("                 ( CD_GB                                                                       ");
                        strSql.AppendLine("                 , COM_CD                                                                      ");
                        strSql.AppendLine("                 , COM_SUB_CD1                                                                 ");
                        strSql.AppendLine("                 , COM_SUB_CD2                                                                 ");
                        strSql.AppendLine("                 , COM_SUB_CD3                                                                 ");
                        strSql.AppendLine("                 , CD_GB_NM                                                                    ");
                        strSql.AppendLine("                 , COM_NM                                                                      ");
                        strSql.AppendLine("                 , COM_SUB_NM1                                                                 ");
                        strSql.AppendLine("                 , COM_SUB_NM2                                                                 ");
                        strSql.AppendLine("                 , COM_SUB_NM3                                                                 ");
                        strSql.AppendLine("                 , CD_AMT1                                                                     ");
                        strSql.AppendLine("                 , CD_AMT2                                                                     ");
                        strSql.AppendLine("                 , CD_AMT3                                                                     ");
                        strSql.AppendLine("                 , CD_YN1                                                                      ");
                        strSql.AppendLine("                 , CD_YN2                                                                      ");
                        strSql.AppendLine("                 , CD_YN3                                                                      ");
                        strSql.AppendLine("                 , SORT_SEQ                                                                    ");
                        strSql.AppendLine("                 , USE_YN                                                                      ");
                        strSql.AppendLine("                 , MFY_DT                                                                      ");
                        strSql.AppendLine("                 , MFY_ID                                                                      ");
                        strSql.AppendLine("                 , REMARK )                                                                    ");
                        strSql.AppendLine("           VALUES('" + sCdGb + "'                                                              ");
                        strSql.AppendLine("                 , '" + sComCd + "'                                                            ");
                        strSql.AppendLine("                 , '" + sComSubCd1 + "'                                                        ");
                        strSql.AppendLine("                 , '" + sComSubCd2 + "'                                                        ");
                        strSql.AppendLine("                 , '" + sComSubCd3 + "'                                                        ");
                        strSql.AppendLine("                 , '" + sCdGbNm + "'                                                           ");
                        strSql.AppendLine("                 , '" + sComNm + "'                                                            ");
                        strSql.AppendLine("                 , '" + sComSubNm1 + "'                                                        ");
                        strSql.AppendLine("                 , '" + sComSubNm2 + "'                                                        ");
                        strSql.AppendLine("                 , '" + sComSubNm3 + "'                                                        ");
                        strSql.AppendLine("                 , '" + sCdAmt1 + "'                                                           ");
                        strSql.AppendLine("                 , '" + sCdAmt2 + "'                                                           ");
                        strSql.AppendLine("                 , '" + sCdAmt3 + "'                                                           ");
                        strSql.AppendLine("                 , '" + sCdYn1 + "'                                                            ");
                        strSql.AppendLine("                 , '" + sCdYn2 + "'                                                            ");
                        strSql.AppendLine("                 , '" + sCdYn3 + "'                                                            ");
                        strSql.AppendLine("                 , '" + sSortSeq + "'                                                          ");
                        strSql.AppendLine("                 , '" + sUseYn + "'                                                            ");
                        strSql.AppendLine("                 , CONVERT(DATE, GETDATE())                                                                   ");
                        strSql.AppendLine("                 , '"+ sUSERCD + "'                                                            ");
                        strSql.AppendLine("                 , '" + sRemark + "'                                                           ");
                        strSql.AppendLine("                )                                                                              ");
                        strSql.AppendLine("      END                                                                                      ");
                                                                                                                                          
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "CD_GB:" + sCdGb + ",CD_GB_NM:" + sCdGbNm + ",COM_CD:" + sComCd + ",COM_NM:" + sComNm;
                        //ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, (j+1).ToString(), "S", this.Name, sLogRmk, cmd);

                        string sIP = ClsFunc.GetLocalIP();

                        strSql.Clear();
                        strSql.AppendLine(" INSERT INTO  ZSYS_LOG ");
                        strSql.AppendLine(" 	         ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                        strSql.AppendLine(" 	    VALUES ");
                        strSql.AppendLine(" 	         ( '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ");
                        strSql.AppendLine(" 	         , '" + FmMainToolBar2.UserID + "'  ");
                        strSql.AppendLine(" 	         , " + (j + 1).ToString() + " ");
                        strSql.AppendLine(" 	         , '" + this.Name + "' ");
                        strSql.AppendLine(" 	         , 'S' ");
                        strSql.AppendLine(" 	         , '" + sIP + "' ");
                        strSql.AppendLine(" 	         , '" + sLogRmk + "' ) ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridCodeDept(sGb);
                }
                else
                {
                    XtraMessageBox.Show("변경된 내용이 없습니다.");
                    int idx = GridViewCodeDept.FocusedRowHandle;

                    GetGridCodeDept(sGb);

                    GridViewCodeDept.FocusedRowHandle = idx;
                }

            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (GridViewCode.GetFocusedRowCellValue("COM_CD") == null) return;
            if ((MessageBox.Show(this, "삭제 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
            {
                return;
            }
            DeleteCode();
        }
        private void DeleteCode()
        {
            int idx = GridViewCode.FocusedRowHandle;

            string sComCd = string.Empty;
            sComCd = GridViewCode.GetFocusedRowCellValue("COM_CD").ToString();
            StringBuilder strSql = new StringBuilder();

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            try
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine("DELETE FROM com_base_cd ");
                strSql.AppendLine(" WHERE CD_GB='COM_COM_CD' ");
                strSql.AppendLine("   AND COM_CD='" + sComCd + "'");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM com_base_cd ");
                strSql.AppendLine("  WHERE CD_GB = '" + sComCd + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "CD_GB:" + sComCd;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm;ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                GetGridTop();

                GridViewCode.FocusedRowHandle = idx - 1;
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
            this.Dispose();
        }

        private void LkupCodeName_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void GridViewCode_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void CommonCdMgt_FormClosed(object sender, FormClosedEventArgs e)
        {
           
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
                string sFileNM = "그룹코드 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GrideCode.ExportToXls(FileName + ".xls");
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

        private void BtnDtlExcel_Click(object sender, EventArgs e)
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
                string sFileNM = "항목코드 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridCodeDept.ExportToXls(FileName + ".xls");
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

        private void CommonCdMgt_TextChanged(object sender, EventArgs e)
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

        private void GridViewCode_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(e.FocusedRowHandle < 0)
            {
                GridCodeDept.DataSource = null;
                return;
            }

            string sGb = GridViewCode.GetFocusedRowCellValue("COM_CD").ToString();
            string sNm = GridViewCode.GetFocusedRowCellValue("COM_NM").ToString();
            GetGb = sGb;
            GetNm = sNm;

            GetGridCodeDept(sGb);
        }

        private void GridViewCode_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName.Equals("COM_CD"))
            {
                string sComcd = GridViewCode.GetRowCellValue(e.RowHandle, e.Column)?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT COUNT(*) AS CNT");
                strSql.AppendLine("	 FROM COM_BASE_CD A");
                strSql.AppendLine(" WHERE CD_GB ='" + sComcd + "'");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sCnt = dt.Rows[0]["CNT"]?.ToString();

                    if(int.TryParse(sCnt, out int iResult))
                    {
                        if(iResult > 0)
                        {
                            XtraMessageBox.Show("이미 사용중인 코드입니다.");
                            GridViewCode.SetRowCellValue(e.RowHandle, e.Column, "");
                            GridViewCode.FocusedColumn = GridColTopComCd;
                        }
                    }
                }
            }
            GridViewCode.UpdateCurrentRow();
        }

        private void GridViewCodeDept_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName.Equals("COM_CD"))
            {
                string sCdgb = GridViewCode.GetFocusedRowCellValue("COM_CD")?.ToString();
                string sComcd = GridViewCodeDept.GetRowCellValue(e.RowHandle, e.Column)?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT COUNT(*) AS CNT");
                strSql.AppendLine("	 FROM COM_BASE_CD A");
                strSql.AppendLine(" WHERE CD_GB ='" + sCdgb + "'");
                strSql.AppendLine("   AND COM_CD = '" + sComcd + "'");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null && dt.Rows.Count > 0)
                {
                    string sCnt = dt.Rows[0]["CNT"]?.ToString();

                    if (int.TryParse(sCnt, out int iResult))
                    {
                        if (iResult > 0)
                        {
                            XtraMessageBox.Show("이미 사용중인 코드입니다.");
                            GridViewCodeDept.SetRowCellValue(e.RowHandle, e.Column, "");
                            GridViewCodeDept.FocusedColumn = GridColComCd;
                        }
                    }
                }
            }
            GridViewCodeDept.UpdateCurrentRow();
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}