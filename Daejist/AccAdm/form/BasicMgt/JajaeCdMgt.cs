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
    public partial class JajaeCdMgt : DevExpress.XtraEditors.XtraForm
    {
        public JajaeCdMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void JajaeCdMgt_Load(object sender, EventArgs e)
        {
            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "JajaeCdMgt");

            DataTable dtUserCd = GetLookUpData("2", "", "", "");
            RepositoryItemGridLookUpEdit userLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GridRetr, GridColUserCd, "CD", "NM", "");

            arrGrdView = new GridView[] { GridViewRetr };
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
                strSql.AppendLine("  WHERE 1=1");

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
            StringBuilder strSql = new StringBuilder();
            strSql.AppendLine("");
            strSql.AppendLine("SELECT J_Spec");
            strSql.AppendLine("     , J_Serial");
            strSql.AppendLine("     , Gubun ");
            strSql.AppendLine("     , MachulGubun");
            strSql.AppendLine("     , DaeGubun ");
            strSql.AppendLine("     , Gubun1");
            strSql.AppendLine("     , Danga ");
            strSql.AppendLine("     , J_Unit");
            strSql.AppendLine("     , SellPrc1 ");
            strSql.AppendLine("     , SellPrc2 ");
            strSql.AppendLine("     , SellPrc3 ");
            strSql.AppendLine("     , SellPrc4");
            strSql.AppendLine("     , SellPrc5 ");
            strSql.AppendLine("     , SellPrc6 ");
            strSql.AppendLine("     , SellPrc7");
            strSql.AppendLine("     , SellPrc8");
            strSql.AppendLine("     , SellPrc9");
            strSql.AppendLine("     , FirstDate");
            strSql.AppendLine("     , LastDate");
            strSql.AppendLine("     , ThickPos ");
            strSql.AppendLine("     , maipgb");
            strSql.AppendLine("     , RMK");
            strSql.AppendLine("  FROM JAJAE");
            strSql.AppendLine(" ORDER BY CASE WHEN RMK = '' OR RMK IS NULL THEN 2 ELSE 1 END, LEN(RMK), RMK");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
            GridRetr.Focus();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if(GridViewRetr.RowCount > 0)
            {
                string sGubun1 =  GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, "Gubun1")?.ToString();

                if (string.IsNullOrEmpty(sGubun1))
                {
                    XtraMessageBox.Show("등급을 입력해주세요.");
                    return;
                }
            }

            GridViewRetr.AddNewRow();
            GridViewRetr.Focus();
            GridViewRetr.FocusedColumn = gridColumn17;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            SaveGrid();
        }
        private void SaveGrid()
        {
            DataTable dt = (DataTable)GridRetr.DataSource;
            DataSet dtSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dtSave.Tables[0];

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string ThickPos = dtMerge.Rows[i]["ThickPos"]?.ToString();

                if (string.IsNullOrEmpty(ThickPos))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            double sJSerial;
            string sGubun = string.Empty;
            string sMachulGubun = string.Empty;
            string sGubun1 = string.Empty;
            //double sDanga = 0;
            string sJUnit = string.Empty;
            //double sSellPrc1 = 0;
            //double sSellPrc2 = 0;
            //double sSellPrc3 = 0;
            //double sSellPrc4 = 0;
            //double sSellPrc5 = 0;
            //double sSellPrc6 = 0;
            //double sSellPrc7 = 0;
            //double sSellPrc8 = 0;
            //double sSellPrc9 = 0;

            string sFirstDate = string.Empty;
            string sLastDate = string.Empty;
            string sDaeGubun = string.Empty;
            double sThickPos = 0;
            string sMaipgb = string.Empty;

            string sRMK = string.Empty;

            
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
                        if (dtMerge.Rows[j]["Gubun"].ToString().Length < 1)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("구분 을 입력해주세요");
                            return;
                        }
                        if (dtMerge.Rows[j]["MachulGubun"].ToString().Length < 1)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("매출구분 을  입력해주세요");
                            return;
                        }
                        if (dtMerge.Rows[j]["DaeGubun"].ToString().Length < 1)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("품목 을 입력해주세요");
                            return;
                        }
                        if (dtMerge.Rows[j]["Gubun1"].ToString().Length < 1)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show("등급명 을 입력해주세요");
                            return;
                        }

                        //if (dtMerge.Rows[j]["Danga"].ToString().Length < 1)
                        //{
                        //    DBConn.dbTran.Rollback();
                        //    DBConn.dbTran = null;
                        //    MessageBox.Show("단가 를 입력해주세요");
                        //    return;
                        //}

                        string sUserCd=FmMainToolBar2.dtUserAutInfo.Rows[0]["USRCD"].ToString();
                        sJSerial = dtMerge.Rows[j]["J_Serial"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["J_Serial"]);
                        sGubun = Convert.ToString(dtMerge.Rows[j]["Gubun"]);
                        sMachulGubun = Convert.ToString(dtMerge.Rows[j]["MachulGubun"]);
                        sGubun1 = Convert.ToString(dtMerge.Rows[j]["Gubun1"]);
                        //sDanga = dtMerge.Rows[j]["Danga"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["Danga"]);
                        sJUnit = Convert.ToString(dtMerge.Rows[j]["J_Unit"]);
                        //sSellPrc1 = dtMerge.Rows[j]["SellPrc1"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc1"]);
                        //sSellPrc2 = dtMerge.Rows[j]["SellPrc2"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc2"]);
                        //sSellPrc3 = dtMerge.Rows[j]["SellPrc3"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc3"]);
                        //sSellPrc4 = dtMerge.Rows[j]["SellPrc4"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc4"]);
                        //sSellPrc5 = dtMerge.Rows[j]["SellPrc5"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc5"]);
                        //sSellPrc6 = dtMerge.Rows[j]["SellPrc6"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc6"]);
                        //sSellPrc7 = dtMerge.Rows[j]["SellPrc7"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc7"]);
                        //sSellPrc8 = dtMerge.Rows[j]["SellPrc8"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc8"]);
                        //sSellPrc9 = dtMerge.Rows[j]["SellPrc9"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["SellPrc9"]);

                        sDaeGubun = Convert.ToString(dtMerge.Rows[j]["DaeGubun"]);
                        sThickPos = dtMerge.Rows[j]["ThickPos"] is DBNull ? 0 : Convert.ToDouble(dtMerge.Rows[j]["ThickPos"]);
                        sMaipgb = dtMerge.Rows[j]["maipgb"]?.ToString();

                        sRMK = dtMerge.Rows[j]["RMK"]?.ToString();

                        if (sJSerial != 0)
                        {
                            //strSql.Clear();
                            //strSql.AppendLine("");
                            //strSql.AppendLine("SELECT ");
                            //strSql.AppendLine("       J_Serial");
                            //strSql.AppendLine("      ,Danga");
                            //strSql.AppendLine("  FROM JAJAE");
                            //strSql.AppendLine(" WHERE J_Serial ='" + sJSerial + "'");
                            //DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            //if (!dt1.Rows[0]["Danga"].Equals(sDanga))
                            //{
                            //    DBConn.dbTran.Rollback();
                            //    DBConn.dbTran = null;
                            //    MessageBox.Show("단가 수정은 불가합니다 ");
                            //    return;
                            //}

                            strSql.Clear();
                            strSql.AppendLine("");
                            strSql.AppendLine(" UPDATE jajae    ");
                            strSql.AppendLine(" SET                 ");
                            strSql.AppendLine("   J_Spec        ='" + sUserCd + "'     ");
                            strSql.AppendLine("  ,RMK        ='" + sRMK + "'     ");
                            strSql.AppendLine("  ,Gubun        ='" + sGubun + "'     ");
                            strSql.AppendLine("  ,MachulGubun  ='" + sMachulGubun + "'     ");
                            strSql.AppendLine("  ,Gubun1       ='" + sGubun1 + "'     ");
                            strSql.AppendLine("  ,J_Unit       ='" + sJUnit + "'     ");
                            strSql.AppendLine("  ,DaeGubun     ='" + sDaeGubun + "'     ");
                            strSql.AppendLine("  ,ThickPos     ='" + sThickPos + "'     ");
                            strSql.AppendLine("  ,maipgb       = '" + sMaipgb + "'");
                            strSql.AppendLine("  ,LastDate     = convert(char(10), getdate(), 23)     ");
                            strSql.AppendLine(" WHERE J_Serial =   '" + sJSerial + "'");
                        }
                        else
                        {
                            strSql.Clear();
                            strSql.AppendLine("");
                            strSql.AppendLine("INSERT INTO JAJAE ");
                            strSql.AppendLine("           (J_Spec      ");
                            //strSql.AppendLine("           ,J_Serial      ");
                            strSql.AppendLine("           ,RMK      ");
                            strSql.AppendLine("           ,Gubun     ");
                            strSql.AppendLine("           ,MachulGubun     ");
                            strSql.AppendLine("           ,Gubun1     ");
                            strSql.AppendLine("           ,J_Unit     ");
                            strSql.AppendLine("           ,FirstDate     ");
                            strSql.AppendLine("           ,DaeGubun     ");
                            strSql.AppendLine("           ,ThickPos     ");
                            strSql.AppendLine("           ,maipgb     ");
                            strSql.AppendLine("           )     ");
                            strSql.AppendLine("     VALUES (");
                            strSql.AppendLine(" 		  '" + sUserCd + "'     ");
                            //strSql.AppendLine("           ,null");
                            strSql.AppendLine("           ,'" + sRMK + "'");
                            strSql.AppendLine(" 		  ,'" + sGubun + "'     ");
                            strSql.AppendLine(" 		  ,'" + sMachulGubun + "'     ");
                            strSql.AppendLine(" 		  ,'" + sGubun1 + "'     ");
                            strSql.AppendLine(" 		  ,'" + sJUnit + "'     ");
                            strSql.AppendLine(" 		  , convert(char(10), getdate(), 23)       ");
                            strSql.AppendLine(" 		  ,'" + sDaeGubun + "'     ");
                            strSql.AppendLine(" 		  ,'" + sThickPos + "'     ");
                            strSql.AppendLine(" 		  ,'" + sMaipgb + "'     ");
                            strSql.AppendLine("            )                    ");
                          

                            //strSql.Clear();
                            //strSql.AppendLine("");
                            //strSql.AppendLine("INSERT INTO JAJAE ");
                            //strSql.AppendLine("           (J_Serial      ");
                            //strSql.AppendLine("           ,Gubun     ");
                            //strSql.AppendLine("           ,MachulGubun     ");
                            //strSql.AppendLine("           ,Gubun1     ");
                            //strSql.AppendLine("           ,Danga     ");
                            //strSql.AppendLine("           ,J_Unit     ");
                            //strSql.AppendLine("           ,SellPrc1     ");
                            //strSql.AppendLine("           ,SellPrc2     ");
                            //strSql.AppendLine("           ,SellPrc4     ");
                            //strSql.AppendLine("           ,SellPrc5     ");
                            //strSql.AppendLine("           ,SellPrc6     ");
                            //strSql.AppendLine("           ,SellPrc7     ");
                            //strSql.AppendLine("           ,FirstDate     ");
                            //strSql.AppendLine("           ,DaeGubun     ");
                            //strSql.AppendLine("           ,ThickPos     ");
                            //strSql.AppendLine("           )     ");
                            //strSql.AppendLine("     VALUES (");                          
                            //strSql.AppendLine("            null");
                            //strSql.AppendLine(" 		  ,'" + sGubun + "'     ");
                            //strSql.AppendLine(" 		  ,'" + sMachulGubun + "'     ");
                            //strSql.AppendLine(" 		  ,'" + sGubun1 + "'     ");
                            //strSql.AppendLine(" 		  ,'" + sDanga + "'     ");
                            //strSql.AppendLine(" 		  ,'" + sJUnit + "'     ");
                            //if (sSellPrc1 != 0) strSql.AppendLine(" 		  ,'" + sSellPrc1 + "'     ");
                            //else strSql.AppendLine("           , 0");
                            //if (sSellPrc2 != 0) strSql.AppendLine(" 		  ,'" + sSellPrc2 + "'     ");
                            //else strSql.AppendLine("           , 0");                           
                            //if (sSellPrc4 != 0) strSql.AppendLine(" 		  ,'" + sSellPrc4 + "'     ");
                            //else strSql.AppendLine("           , 0");
                            //if (sSellPrc5 != 0) strSql.AppendLine(" 		  ,'" + sSellPrc5 + "'     ");
                            //else strSql.AppendLine("           , 0");
                            //if (sSellPrc6 != 0) strSql.AppendLine(" 		  ,'" + sSellPrc6 + "'     ");
                            //else strSql.AppendLine("           , 0");
                            //if (sSellPrc7 != 0) strSql.AppendLine(" 		  ,'" + sSellPrc7 + "'     ");
                            //else strSql.AppendLine("           , 0");
                            //strSql.AppendLine(" 		  , NOW()       ");
                            //strSql.AppendLine(" 		  ,'" + sDaeGubun + "'     ");
                            //strSql.AppendLine(" 		  ,'" + sThickPos + "'     ");
                            //strSql.AppendLine("            )                    ");

                        }

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "Table:jajae -> J_Spec:" + sUserCd + "Gubun1" + sGubun1;
                        ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, (j+1).ToString(), "S", this.Name, sLogRmk, cmd);
                    }
                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridRetr();

                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, gridColumn5, dtMerge.Rows[0]["Gubun1"]?.ToString());
                
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
                XtraMessageBox.Show("변경된 내용이 없습니다.");
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void JajaeCdMgt_KeyDown(object sender, KeyEventArgs e)
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

        private void JajaeCdMgt_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sJSerial = GridViewRetr.GetFocusedRowCellValue("J_Serial")?.ToString();
            string sGubun1 = GridViewRetr.GetFocusedRowCellValue("Gubun1")?.ToString();

            if (string.IsNullOrEmpty(sJSerial))
            {
                XtraMessageBox.Show("자재코드가 존재하지 않습니다.\r\n삭제하려는 데이터를 정확히 클릭하세요.");
                return;
            }

            if (XtraMessageBox.Show("자재코드 : " + sJSerial + "\r\n자재명 : " + sGubun1 + " \r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
              , "자재코드 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
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
                strSql.AppendLine(" DELETE FROM JAJAE ");
                strSql.AppendLine("       WHERE J_SERIAL = '" + sJSerial + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:JAJAE -> J_SERIAL:" + sJSerial;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = idx-1;
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
                string sFileNM = "자재코드 리스트";
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

        private void JajaeCdMgt_TextChanged(object sender, EventArgs e)
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

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr.UpdateCurrentRow();
        }
    }
}