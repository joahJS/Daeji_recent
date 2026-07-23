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
using System.IO;
using System.Data.SqlClient;
using DevExpress.XtraSplashScreen;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.XtraGrid.Views.BandedGrid;

namespace AccAdm
{
    public partial class IN11001F01 : DevExpress.XtraEditors.XtraForm
    {
        private string PROCEDURE_ID = "DP_AllReport";
        string stMagam;

        public IN11001F01()
        {
            InitializeComponent();
        }

        private void IN11001F01_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            DateFrom.EditValue = DateTime.Today;
            ComnEtcFunc.SetDateFromToValue(DateFrom, DateTo);
            DateTo.EditValue = DateTime.Today;
            DateYm.EditValue = DateTime.Today;

            DataTable dtGB = GetLookUpData("1", "", "N");

            ComLib.ComGrid.SetGridLookUpEdit(RepoGridLkupGubun, dtGB, GridRetr1, GridColGubun1, "CD", "NM", "");
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr1 };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
        }
        #endregion

        #region lookup
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

            if (sGb.Equals("1")) 
            {
                strSql.AppendLine(" SELECT J_SERIAL AS CD       ");
                strSql.AppendLine("      , GUBUN1 AS NM         ");
                strSql.AppendLine("   FROM JAJAE                ");
                strSql.AppendLine("  WHERE DAEGUBUN = '재고이동'");
                strSql.AppendLine("  UNION ALL");
                strSql.AppendLine(" SELECT J_SERIAL AS CD    ");
                strSql.AppendLine("      , GUBUN1 AS NM      ");
                strSql.AppendLine("   FROM JAJAE             ");
                strSql.AppendLine("  WHERE J_SERIAL IN (2025163,650,651,652,653)");
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
        #endregion


        #region [일마감]
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void BtnRetr1_Click(object sender, EventArgs e)
        {
            string sDateFrom = DateFrom.EditValue?.ToString().Substring(0,10);
            string sDateTo = DateTo.EditValue?.ToString().Substring(0, 10);
            string sChk = RadiMagm.EditValue?.ToString();
            string sGB = RadIGB.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A1.JUNPYOID                                                                      ");
            strSql.AppendLine("      , A1.J_DATE                                                                      ");
            strSql.AppendLine("      , A1.KERATYPE                                                                    ");
            strSql.AppendLine("      , A1.J_CHECK                                                                     ");
            strSql.AppendLine("      , CASE WHEN A1.KERATYPE = '입고' THEN A1.MAIPCHER ELSE A1.J_COMPANY END AS CVNAM ");
            strSql.AppendLine("      , A1.J_SERIAL AS GUBUN1                                                          ");
            strSql.AppendLine("      , A1.GUMSUBIGO AS RK                                                             ");
            strSql.AppendLine("      , CASE WHEN A1.KERATYPE = '입고' THEN A1.IWEIGHT ELSE A1.OWEIGHT END AS WEIGHT                                                                      ");
            strSql.AppendLine("      , CASE WHEN A1.KERATYPE = '입고' THEN A1.IDANGA ELSE A1.ODANGA END AS DANGA      ");
            strSql.AppendLine("      , CASE WHEN A1.KERATYPE = '입고' THEN A1.IKONGKEP ELSE A1.OKONGKEP END AS AMT    ");
            strSql.AppendLine("   FROM MESURING A1                                                                    ");
            strSql.AppendLine("   LEFT JOIN JAJAE B1                                                                  ");
            strSql.AppendLine("     ON A1.J_SERIAL = B1.J_SERIAL                                                      ");
            //strSql.AppendLine("  WHERE B1.DAEGUBUN = '재고이동'                                                       ");
            strSql.AppendLine("  WHERE 1 = 1                                                                             ");
            strSql.AppendLine("    AND A1.J_Date BETWEEN '"+ sDateFrom + "' AND '"+ sDateTo + "'                            ");
            //strSql.AppendLine("    AND A1.KeraType = '출고'                             ");
            if (!sChk.Equals("*"))
            {
                strSql.AppendLine("AND A1.J_CHECK = '"+ sChk + "'");
            }
            if (sGB.Equals("재고이동"))
            {
                strSql.AppendLine("AND (B1.DAEGUBUN = '재고이동' OR A1.MaipCher = '재고이동' OR A1.J_Company = '재고이동')");
            }else if (sGB.Equals("ASR"))
            {
                strSql.AppendLine("AND A1.J_SERIAL = '2025163'");
            }
            else if (sGB.Equals("비철"))
            {
                strSql.AppendLine("AND (B1.DAEGUBUN = '슈레더' AND A1.J_SERIAL IN (650,651,652,653))");
            }
            else
            {
                strSql.AppendLine("    AND (B1.DAEGUBUN = '재고이동' OR A1.MaipCher = '재고이동' OR A1.J_Company = '재고이동' OR A1.J_SERIAL = '2025163' OR (B1.DAEGUBUN = '슈레더' AND A1.J_SERIAL IN (650,651,652,653)))");//ASR 추가, 2022-03-02 비철추가
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt != null && dt.Rows.Count > 0)
            {
                GridViewRetr1.Focus();
                GridViewRetr1.FocusedColumn = GridColDanga;
            }
            else
            {
                DateFrom.Focus();
            }

            GridRetr1.DataSource = dt;
        }

        private void GridViewRetr1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void RadiMagm_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr1.PerformClick();

            string sMagm = RadiMagm.EditValue?.ToString();
            if (sMagm.Equals("1"))
            {
                BtnDayMG.Enabled = true;
                BtnDayMG.Text = "마감취소(F3)";

                BtnDanAll.Enabled = false;

                GridColGubun1.OptionsColumn.AllowEdit = false;
                //GridColWeight.OptionsColumn.AllowEdit = false;
                GridColDanga.OptionsColumn.AllowEdit = false;
                GridColAmt.OptionsColumn.AllowEdit = false;
            }
            else if(sMagm.Equals(""))
            {
                BtnDayMG.Enabled = true;
                BtnDayMG.Text = "일 마감(F3)";

                BtnDanAll.Enabled = true;

                GridColGubun1.OptionsColumn.AllowEdit = true;
                //GridColWeight.OptionsColumn.AllowEdit = true;
                GridColDanga.OptionsColumn.AllowEdit = true;
                GridColAmt.OptionsColumn.AllowEdit = true;
            }
            else
            {
                BtnDayMG.Enabled = false;

                BtnDanAll.Enabled = false;

                GridColGubun1.OptionsColumn.AllowEdit = false;
                //GridColWeight.OptionsColumn.AllowEdit = false;
                GridColDanga.OptionsColumn.AllowEdit = false;
                GridColAmt.OptionsColumn.AllowEdit = false;
            }
        }

        private void RadiIpChul_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr1.PerformClick();
        }

        private void DateYmd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr1.PerformClick();
        }

        private void GridViewRetr1_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == GridColDanga /*|| e.Column == GridColAmt*/)
            {
                double dVal = string.IsNullOrEmpty(e.CellValue?.ToString()) ? 0 : Convert.ToDouble(e.CellValue?.ToString());
                if (dVal > 0)
                    e.Appearance.BackColor = Color.PaleGreen;
                else if (dVal == 0)
                    e.Appearance.BackColor = SystemColors.Info;
            }
        }

        private void GridViewRetr1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr1.SelectRow(e.RowHandle);
        }

        private void BtnDayMG_Click(object sender, EventArgs e)
        {
            string sText = BtnDayMG.Text;

            if (sText.Equals("마감취소(F3)"))
            {
                CancelMaGam();
            }
            else if (sText.Equals("일 마감(F3)"))
            {
                SaveMaGam();
            }
        }

        //일마감
        private void SaveMaGam()
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int[] selectIndex = GridViewRetr1.GetSelectedRows();

            if (selectIndex.Length == 0)
            {
                XtraMessageBox.Show("마감할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show("선택한" + selectIndex.Length + "개의 항목에 대해 일마감을 진행하시겠습니까?", "마감여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StringBuilder strSql = new StringBuilder();

            DataTable dtRetr = GridRetr1.DataSource as DataTable;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                for (int i = 0; i < selectIndex.Length; i++)
                {
                    //정렬 후 변경된 index값 원래 index값으로 변경
                    int orignalIndex = GridViewRetr1.GetDataSourceRowIndex(selectIndex[i]);

                    DataRow row = dtRetr.Rows[orignalIndex];

                    string sJunpyoID = row["JUNPYOID"]?.ToString();
                    string sJ_Serial = row["GUBUN1"]?.ToString();
                    string sDanGa = row["DANGA"]?.ToString();
                    string sWeight = row["WEIGHT"]?.ToString();
                    string sKERATYPE = row["KERATYPE"]?.ToString();

                    string sBasdt = row["J_DATE"]?.ToString();
                    string sPrice = row["AMT"]?.ToString();
                    string sUser = FmMainToolBar2.UserID;

                    //계근정보 수정(마감처리, 단가, 중량수정(X))
                    strSql.Clear();
                    strSql.AppendLine(" DECLARE @GUBUN1 VARCHAR(25);");
                    strSql.AppendLine("                             ");
                    strSql.AppendLine(" SELECT @GUBUN1 = GUBUN1     ");
                    strSql.AppendLine("   FROM JAJAE                ");
                    strSql.AppendLine("  WHERE J_SERIAL = " + sJ_Serial + "         ");
                    strSql.AppendLine("                             ");
                    strSql.AppendLine(" UPDATE MESURING             ");
                    strSql.AppendLine("    SET J_SERIAL = " + sJ_Serial + "         ");
                    strSql.AppendLine("      , J_CHECK = '1'         ");

                    if (sKERATYPE.Equals("입고"))
                    {
                        strSql.AppendLine("      , IDANGA = " + sDanGa + "          ");
                        strSql.AppendLine("      , IKONGKEP = " + sPrice + "          ");
                    }
                    else
                    {
                        strSql.AppendLine("      , ODANGA = " + sDanGa + "          ");
                        strSql.AppendLine("      , OKONGKEP = " + sPrice + "          ");
                    }

                    //strSql.AppendLine("      , WEIGHT = " + sWeight + "          ");
                    strSql.AppendLine("      , GUBUN1 = @GUBUN1         ");
                    strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoID + "         ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    //재고 일마감
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT SLINO FROM J_MAGAM WHERE M_ID = "+ sJunpyoID + ")        ");
                    strSql.AppendLine("    BEGIN                                                    ");
                    strSql.AppendLine("          UPDATE J_MAGAM                                     ");
                    strSql.AppendLine("             SET DANGA = "+ sDanGa + "                                   ");
	                strSql.AppendLine("               , WEIGHT = "+ sWeight + "                                  ");
	                strSql.AppendLine("               , J_SERIAL = "+ sJ_Serial + "                                ");
	                strSql.AppendLine("               , PRICE = "+ sPrice + "                                   ");
	                strSql.AppendLine("               , MDATE = CONVERT(VARCHAR(25), GETDATE(), 20) ");
	                strSql.AppendLine("               , MUSER = "+ sUser + "                                   ");
                    strSql.AppendLine("           WHERE M_ID = "+ sJunpyoID + "                                   ");
                    strSql.AppendLine("      END                                                    ");
                    strSql.AppendLine(" ELSE                                                        ");
                    strSql.AppendLine("    BEGIN                                                    ");
                    strSql.AppendLine("          DECLARE @SLINO NUMERIC                             ");
                    strSql.AppendLine("                                                             ");
                    strSql.AppendLine("          SELECT @SLINO = ISNULL(MAX(SLINO), 0) + 1          ");
                    strSql.AppendLine("            FROM J_MAGAM                                     ");
                    strSql.AppendLine("                                                             ");
                    strSql.AppendLine("          INSERT INTO J_MAGAM( SLINO                         ");
                    strSql.AppendLine("                             , M_ID                          ");
                    strSql.AppendLine("                             , BASDT                         ");
                    strSql.AppendLine("                             , J_SERIAL                      ");
                    strSql.AppendLine("                             , WEIGHT                        ");
                    strSql.AppendLine("                             , DANGA                         ");
                    strSql.AppendLine("                             , PRICE                         ");
                    strSql.AppendLine("                             , CUSER)                        ");
                    strSql.AppendLine("                       VALUES(@SLINO                         ");
                    strSql.AppendLine("                             , "+ sJunpyoID + "                             ");
                    strSql.AppendLine("                             , '"+ sBasdt + "'                             ");
                    strSql.AppendLine("                             , "+ sJ_Serial + "                             ");
                    strSql.AppendLine("                             , "+ sWeight + "                             ");
                    strSql.AppendLine("                             , "+ sDanGa + "                             ");
                    strSql.AppendLine("                             , "+ sPrice + "                             ");
                    strSql.AppendLine("                             , "+ sUser + ")                            ");
                    strSql.AppendLine("      END                                                    ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("일마감을 완료했습니다.");

                BtnRetr1.PerformClick();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        //일마감 취소
        private void CancelMaGam()
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int[] selectIndex = GridViewRetr1.GetSelectedRows();

            if (selectIndex.Length == 0)
            {
                XtraMessageBox.Show("마감취소 할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show("선택한" + selectIndex.Length + "개의 항목에 대해 마감취소를 진행하시겠습니까?", "마감여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StringBuilder strSql = new StringBuilder();

            DataTable dtRetr = GridRetr1.DataSource as DataTable;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                for (int i = 0; i < selectIndex.Length; i++)
                {
                    //정렬 후 변경된 index값 원래 index값으로 변경
                    int orignalIndex = GridViewRetr1.GetDataSourceRowIndex(selectIndex[i]);

                    DataRow row = dtRetr.Rows[orignalIndex];

                    string sJunpyoID = row["JUNPYOID"]?.ToString();
                    string sJ_Serial = row["GUBUN1"]?.ToString();
                    string sDanGa = row["DANGA"]?.ToString();
                    string sWeight = row["WEIGHT"]?.ToString();
                    string sKERATYPE = row["KERATYPE"]?.ToString();

                    string sBasdt = row["J_DATE"]?.ToString();
                    string sPrice = row["AMT"]?.ToString();
                    string sUser = FmMainToolBar2.UserID;

                    //계근정보 수정(마감취소 처리)
                    strSql.Clear();
                    strSql.AppendLine(" UPDATE MESURING             ");
                    strSql.AppendLine("    SET J_CHECK = ''         ");
                    strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoID + "         ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    //재고 일마감 삭제
                    strSql.Clear();
                    strSql.AppendLine(" DELETE FROM J_MAGAM");
                    strSql.AppendLine("  WHERE M_ID = "+ sJunpyoID );

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("마감취소를 완료했습니다.");

                BtnRetr1.PerformClick();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void RepoTxtDanga_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txtedit = (TextEdit)sender;

            string sWeight = GridViewRetr1.GetFocusedRowCellValue(GridColWeight)?.ToString();
            string sDanga = txtedit.EditValue?.ToString();

            if (string.IsNullOrEmpty(sDanga) || string.IsNullOrEmpty(sWeight))
                return;

            double dAmt = double.Parse(sWeight) * double.Parse(sDanga);

            GridViewRetr1.SetFocusedRowCellValue(GridColAmt, dAmt);
        }

        private void RepoTxtWeight_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txtedit = (TextEdit)sender;

            string sWeight = txtedit.EditValue?.ToString();
            string sDanga = GridViewRetr1.GetFocusedRowCellValue(GridColDanga)?.ToString();

            if (string.IsNullOrEmpty(sDanga) || string.IsNullOrEmpty(sWeight))
                return;

            double dAmt = double.Parse(sWeight) * double.Parse(sDanga);

            GridViewRetr1.SetFocusedRowCellValue(GridColAmt, dAmt);
        }
        #endregion

        #region [월마감]
        private void BtnRetr2_Click(object sender, EventArgs e)
        {
            int year = Convert.ToInt16(DateYm.EditValue?.ToString().Substring(0, 4));
            int month = Convert.ToInt16(DateYm.EditValue?.ToString().Substring(5, 2));
            if (DateYm.EditValue?.ToString().Substring(0, 10) == DateYm.EditValue?.ToString().Substring(0, 8) + Convert.ToString(DateTime.DaysInMonth(year, month)))
            {
                layoutControlItem19.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                BtnYmMagam.Visible = true;
            }
            else
            {
                layoutControlItem19.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                BtnYmMagam.Visible = false;
            }

            SplashScreenManager.ShowForm(typeof(WaitForm1));

            GetJaeGoData();

            SplashScreenManager.CloseForm();
        }

        private void GetJaeGoData()
        {
            string sFrom = DateYm.EditValue?.ToString().Substring(0, 10);
            string sTo = DateYm.EditValue?.ToString().Substring(0, 10);
            DateTime tDateTime = Convert.ToDateTime(DateYm.EditValue?.ToString().Substring(0, 10));
            stMagam = Convert.ToString(tDateTime.AddMonths(1))?.Substring(0, 8) + "01";



            string stFrom = Convert.ToString(tDateTime.ToString())?.Substring(0, 8) + "01";

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST4");
            dicParams.Add("SDATE", stFrom);

            DataTable dtC = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            if (dtC.Rows.Count == 0)
            {
                GridRetr1.DataSource = null;
                XtraMessageBox.Show("'" + stFrom + "'월 마감이 되어있지 않습니다. 마감 후 조회해주세요");
                return;
            }

            dicParams.Clear();
            dicParams.Add("CMD", "LIST1");
            dicParams.Add("FDATE", sFrom);
            dicParams.Add("TDATE", sTo);

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

            GridRetr2.DataSource = dt;
        }




        private void BtnYmMagam_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string InsertSql, DelSql;

            DelSql = "DELETE FROM ARBASIC WHERE TDATE = '" + stMagam + "'  ";
            SqlCommand mwDel = new SqlCommand(DelSql, DBConn.dbCon);
            mwDel.ExecuteNonQuery();



            InsertSql = " INSERT INTO  ARBASIC VALUES (@TDATE, @GOCULW, @GOCULD, @GOCULC, @SURDW, @SURDD, @SURDC, @TOTALW, @TOTALD, @TOTALC ) ";
            SqlCommand mwInsert = new SqlCommand(InsertSql, DBConn.dbCon);
            mwInsert.Parameters.Add("@TDATE", SqlDbType.VarChar).Value = stMagam;
            mwInsert.Parameters.Add("@GOCULW", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "A2").ToString());
            mwInsert.Parameters.Add("@GOCULD", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "B2").ToString());
            mwInsert.Parameters.Add("@GOCULC", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "C2").ToString());
            mwInsert.Parameters.Add("@SURDW", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "A1").ToString());
            mwInsert.Parameters.Add("@SURDD", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "B1").ToString());
            mwInsert.Parameters.Add("@SURDC", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "C1").ToString());
            mwInsert.Parameters.Add("@TOTALW", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "A3").ToString());
            mwInsert.Parameters.Add("@TOTALD", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "B3").ToString());
            mwInsert.Parameters.Add("@TOTALC", SqlDbType.VarChar).Value = Convert.ToDecimal(GridViewRetr2.GetRowCellDisplayText(7, "C3").ToString());

            mwInsert.ExecuteNonQuery();
            Cursor = Cursors.Default;
            XtraMessageBox.Show("월마감이 완료되었습니다.");
        }

        

        private void BtnExcel2_Click(object sender, EventArgs e)
        {
            FileInfo_1 fileInfo = new FileInfo_1("4");

            Cursor = Cursors.WaitCursor;
            string[] sPath = fileInfo.CheckFileInfo();
            Cursor = Cursors.Default;

            if (sPath != null)
            {
                SetWeeklyReportForm(sPath[0], sPath[1]);
            }
        }

        Excel.Application ExcelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;
        private void SetWeeklyReportForm(string StandardPath, string SavePath)
        {
            try
            {
                if (!File.Exists(StandardPath))
                {
                    XtraMessageBox.Show("엑셀파일 양식이 존재하지 않습니다.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                string sDate = DateYm.EditValue?.ToString().Substring(0, 10);

                ExcelApp = new Excel.Application();
                wb = ExcelApp.Workbooks.Open(StandardPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                ws = wb.Worksheets.get_Item(1);

                ws.Name = sDate.Substring(5, 2) + "월";

                DataTable dt = (DataTable) GridRetr2.DataSource;

                if(dt == null)
                {
                    XtraMessageBox.Show("엑셀로 내려받을 데이터가 없습니다.\r\n조회 후 이용해주세요.");
                    return;
                }

                //일자부분 세팅
                ws.Range["B3"].Value = DateTime.Parse(sDate).ToString("yyyy년 MM월");
                int iStartRow = 5;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sWEIGHT1  = dt.Rows[i]["A1"]?.ToString();
                    string sDANGA1   = dt.Rows[i]["B1"]?.ToString();
                    string sKONGKEP1 = dt.Rows[i]["C1"]?.ToString();
                    string sWEIGHT2  = dt.Rows[i]["A2"]?.ToString();
                    string sDANGA2   = dt.Rows[i]["B2"]?.ToString();
                    string sKONGKEP2 = dt.Rows[i]["C2"]?.ToString();
                    
                    string sWEIGHT6  = dt.Rows[i]["A3"]?.ToString();
                    string sDANGA6   = dt.Rows[i]["B3"]?.ToString();
                    string sKONGKEP6 = dt.Rows[i]["C3"]?.ToString();

                    int iApplyRowIdx = iStartRow + (i + 1);

                    ws.Range["C" + iApplyRowIdx].Value = sWEIGHT1;
                    ws.Range["D" + iApplyRowIdx].Value = sDANGA1;
                    ws.Range["E" + iApplyRowIdx].Value = sKONGKEP1;
                    ws.Range["F" + iApplyRowIdx].Value = sWEIGHT2;
                    ws.Range["G" + iApplyRowIdx].Value = sDANGA2;
                    ws.Range["H" + iApplyRowIdx].Value = sKONGKEP2;
                   
                    ws.Range["I" + iApplyRowIdx].Value = sWEIGHT6;
                    ws.Range["J" + iApplyRowIdx].Value = sDANGA6;
                    ws.Range["K" + iApplyRowIdx].Value = sKONGKEP6;
                }

                if (File.Exists(SavePath))
                    File.Delete(SavePath);

                Cursor = Cursors.Default;
                //wb.SaveAs(SavePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //파일 닫기... 
                wb.SaveAs(SavePath);
                wb.Close(false, Type.Missing, Type.Missing);
                wb = null;
                ExcelApp.Quit();

                Process.Start(SavePath);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                XtraMessageBox.Show(ex.Message);

                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
            finally
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();
            }
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        private void IN11001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage.Name.Equals("XTabDay"))
            {
                if (e.KeyCode == Keys.Escape) { }
                else if (e.KeyCode == Keys.F3)
                    BtnDayMG.PerformClick();
                else if (e.KeyCode == Keys.F5)
                    BtnRetr1.PerformClick();
                else if (e.KeyCode == Keys.F8)
                    BtnExcel1.PerformClick();
            }
            else if (xtraTabControl1.SelectedTabPage.Name.Equals("XTabYM"))
            {
                if (e.KeyCode == Keys.Escape) { }
                else if (e.KeyCode == Keys.F3)
                    BtnYmMagam.PerformClick();
                else if (e.KeyCode == Keys.F5)
                    BtnRetr2.PerformClick();
                else if (e.KeyCode == Keys.F8)
                    BtnExcel2.PerformClick();
            }
        }

        private void xtraTabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage.Name.Equals("XTabDay"))
            {
                DateFrom.Focus();
            }
            else if (xtraTabControl1.SelectedTabPage.Name.Equals("XTabYM"))
            {
                DateYm.Focus();
            }
        }

        private void DateYm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr2.PerformClick();
        }

        private void BGridViewYMRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            string title = GridViewRetr2.GetRowCellValue(e.RowHandle, "GU")?.ToString();

            if (DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName != "DevExpress Dark Style")
            {
                if (!string.IsNullOrEmpty(title) && (title.Equals("월초재고") || title.Equals("월말재고")))
                {
                    e.Appearance.BackColor = Color.LightGreen;
                }
                else if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }
            }
            else if (DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName == "DevExpress Dark Style")
            {
                if (!string.IsNullOrEmpty(title) && (title.Equals("월초재고") || title.Equals("월말재고")))
                {
                    e.Appearance.BackColor = Color.LightGreen;
                }
                else if(e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.ControlDark;
                }
            }
        }

        string _Danga = string.Empty;
        private void BtnDanAll_Click(object sender, EventArgs e)
        {
            int[] selectIndex = GridViewRetr1.GetSelectedRows();

            if (selectIndex.Length == 0)
            {
                XtraMessageBox.Show("단가 일괄 등록할 항목을 선택해주세요.");
                return;
            }

            IN11001F02 frm = new IN11001F02();

            frm.Owner = this;
            frm.DataRowSendEvent += new IN11001F02.SendDataHandler(SetDanGaAll);

            if (frm.ShowDialog() == DialogResult.OK)
            {
                double dDanga = 0;

                double.TryParse(_Danga, out dDanga);

                for(int i=0;i< selectIndex.Length; i++)
                {
                    double dWeight = 0;

                    double.TryParse(GridViewRetr1.GetRowCellValue(selectIndex[i], GridColWeight)?.ToString(), out dWeight);

                    GridViewRetr1.SetRowCellValue(selectIndex[i], GridColDanga, dDanga);
                    GridViewRetr1.SetRowCellValue(selectIndex[i], GridColAmt, dWeight * dDanga);
                }
            }
        }
        
        private void SetDanGaAll(DataRow row)
        {
            _Danga = row["DANGA"]?.ToString();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateYm.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateYm.EditValue = sPrevDate;

                BtnRetr2.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateYm.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateYm.EditValue = sNextDate;

                BtnRetr2.PerformClick();
            }
        }

        private void IN11001F01_TextChanged(object sender, EventArgs e)
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

        private void GridViewRetr2_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string sVal = GridViewRetr2.GetRowCellValue(e.RowHandle, "GU")?.ToString();

            if (!string.IsNullOrEmpty(sVal) && sVal.Equals("소계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
                e.Appearance.ForeColor = Color.Black;
            }
            else if (!string.IsNullOrEmpty(sVal) && sVal.Equals("합계"))
            {
                e.Appearance.BackColor = Color.PaleGreen;
                e.Appearance.ForeColor = Color.Black;
            }
            else
            {
                BandedGridView gridView = sender as BandedGridView;

                if (e.RowHandle == gridView.FocusedRowHandle)
                {
                    // 클릭된 행의 배경색 변경
                    e.Appearance.BackColor = Color.FromArgb(230, 230, 230);
                    e.Appearance.ForeColor = Color.Black;
                }
                else
                {
                    // 클릭되지 않은 다른 행의 배경색을 기본으로 설정
                    e.Appearance.BackColor = SystemColors.Window;
                }
            }
        }

        static double sumWeight, sumAmt = 0;
        private void BtnExcel1_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //}
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Add();
                ws = wb.ActiveSheet;

                ws.Cells[1, 2] = "조회일자";
                ws.Cells[1, 3] = DateFrom.DateTime.ToString("yyyy-MM-dd");
                ws.Cells[1, 4] = DateTo.DateTime.ToString("yyyy-MM-dd");

                for (int col = 0; col < GridViewRetr1.Columns.Count; col++)
                {
                    ws.Cells[2, col + 2] = GridViewRetr1.Columns[col].Caption;
                }

                for (int i = 0; i < GridViewRetr1.RowCount; i++) 
                {
                    ws.Cells[i + 3, 2] = GridViewRetr1.GetRowCellValue(i, GridColJdate);
                    ws.Cells[i + 3, 3] = GridViewRetr1.GetRowCellValue(i, GridColKeratype);
                    ws.Cells[i + 3, 4] = GridViewRetr1.GetRowCellValue(i, GridColJChk);
                    ws.Cells[i + 3, 5] = GridViewRetr1.GetRowCellValue(i, GridColCvnam);
                    ws.Cells[i + 3, 6] = GridViewRetr1.GetRowCellValue(i, gridColumn1);
                    ws.Cells[i + 3, 7] = GridViewRetr1.GetRowCellDisplayText(i, GridColGubun1);
                    ws.Cells[i + 3, 8] = GridViewRetr1.GetRowCellValue(i, GridColWeight);
                    ws.Cells[i + 3, 9] = GridViewRetr1.GetRowCellValue(i, GridColDanga);
                    ws.Cells[i + 3, 10] = GridViewRetr1.GetRowCellValue(i, GridColAmt);
                }

                for (int row = 0; row < GridViewRetr1.RowCount; row++)
                {
                    sumWeight += Convert.ToDouble(GridViewRetr1.GetRowCellValue(row, GridColWeight));
                    sumAmt += Convert.ToDouble(GridViewRetr1.GetRowCellValue(row, GridColAmt));
                }
                ws.Cells[GridViewRetr1.RowCount + 3, 8] = sumWeight;
                ws.Cells[GridViewRetr1.RowCount + 3, 10] = sumAmt;


                excelApp.Visible = true;
            }
            finally
            {

            }
        }
    }
}