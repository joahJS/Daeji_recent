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
using System.IO;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class EquipToolMgt : DevExpress.XtraEditors.XtraForm
    {
        public EquipToolMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void ToolManagementDev_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.gp_SetColorFocused(layoutControl3);

            SetLookUpEdit(LkupEditDept, "1", "Y", "Y");

            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            
            arrGrdView = new GridView[] { GridViewToolManagement, GridViewToolRecord };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            DataTable dt = GetLookUpData();
            ComGrid.SetLookUpEdit(LkupToolNmRetr, dt, "CD", "CD", "Y");
            LkupToolNmRetr.Properties.PopulateColumns();
            LkupToolNmRetr.Properties.Columns[1].Visible = false;

            BtnRetr_Click(null, null);
        }

        private DataTable GetLookUpData()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");
            strSql.AppendLine(" SELECT '' AS CD");
            strSql.AppendLine("     , '' AS NM");
            strSql.AppendLine(" UNION ALL");

            strSql.AppendLine(" SELECT TOOL_NM AS CD ");
            strSql.AppendLine(" 	 , '' AS NM      ");
            strSql.AppendLine(" FROM EQUIP_TOOL_MGT  ");
            strSql.AppendLine(" WHERE DELETE_YN = '"+ RdgbMenuDeleteYn.EditValue.ToString() + "'");
            strSql.AppendLine(" GROUP BY TOOL_NM     ");

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }

            StringBuilder strSql = new StringBuilder();
            string sToolNm = LkupToolNmRetr.Text;
            string sDeleteYn = RdgbMenuDeleteYn.EditValue.ToString();

            Cursor = Cursors.WaitCursor;

            strSql.AppendLine(" SELECT A.MG_NO ");
            strSql.AppendLine("      , A.TOOL_NM");
            strSql.AppendLine("      , A.MG_DEPT ");
            strSql.AppendLine("      , A.MODEL_NM ");
            strSql.AppendLine("      , A.STANDARD");
            strSql.AppendLine("      , A.TOOL_PRC ");
            strSql.AppendLine("      , A.MADE_COMPANY");
            strSql.AppendLine("      , A.PUR_DATE ");
            strSql.AppendLine("      , A.DELETE_YN ");
            strSql.AppendLine("      , A.DELETE_REASON ");
            strSql.AppendLine("      , A.USE_COMPO ");
            strSql.AppendLine("      , A.MANAGER1 ");
            strSql.AppendLine("      , A.MANAGER2 ");
            strSql.AppendLine("      , A.TOOL_IMG");
            strSql.AppendLine("   FROM EQUIP_TOOL_MGT A ");
            strSql.AppendLine("  WHERE 1=1");
            strSql.AppendLine("    AND A.DELETE_YN = '" + sDeleteYn + "' ");
            if (!string.IsNullOrEmpty(sToolNm))
            {
                strSql.AppendLine("    AND A.TOOL_NM = '" + sToolNm + "' ");
            }
            strSql.AppendLine("  ORDER BY A.MG_No ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridToolManagement.DataSource = dt;

            //GridToolManagement.RestoreLayoutFromXml(@"C:\STLNT\AccAdm\xaml\ToolManagementDev.xaml");

            Cursor = Cursors.Default;
        }
        private void SetLookUpEdit(DevExpress.XtraEditors.LookUpEdit lkup, string sGb, string sNullYn, string sSetIdx)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD ");
                strSql.AppendLine("      , '전체' AS NM");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");

            }

            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            lkup.Properties.DataSource = dt;
            lkup.Properties.DisplayMember = "NM";
            lkup.Properties.ValueMember = "CD";

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }
        
        private void GridViewToolManagement_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            Byte[] ImgData;
            string sMg_No = GridViewToolManagement.GetFocusedRowCellValue("MG_NO") == null ? string.Empty : GridViewToolManagement.GetFocusedRowCellValue("MG_NO").ToString();
            TxtToolNm.Text = GridViewToolManagement.GetFocusedRowCellValue("TOOL_NM") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("TOOL_NM").ToString();
            LkupEditDept.Text = GridViewToolManagement.GetFocusedRowCellValue("MG_DEPT")== null ? "" : GridViewToolManagement.GetFocusedRowCellValue("MG_DEPT").ToString();
            TxtModelNm.Text = GridViewToolManagement.GetFocusedRowCellValue("MODEL_NM") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("MODEL_NM").ToString();
            TxtMgNo.Text = GridViewToolManagement.GetFocusedRowCellValue("MG_NO") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("MG_NO").ToString();
            TxtStandard.Text = GridViewToolManagement.GetFocusedRowCellValue("STANDARD") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("STANDARD").ToString();
            TxtToolPrc.Text = GridViewToolManagement.GetFocusedRowCellValue("TOOL_PRC") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("TOOL_PRC").ToString();
            TxtMadeCompany.Text = GridViewToolManagement.GetFocusedRowCellValue("MADE_COMPANY") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("MADE_COMPANY").ToString();
            DateEditPurDate.EditValue = GridViewToolManagement.GetFocusedRowCellValue("PUR_DATE") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("PUR_DATE").ToString();
            RdgbDeleteYn.EditValue = GridViewToolManagement.GetFocusedRowCellValue("DELETE_YN") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("DELETE_YN").ToString();
            TxtDeleteReason.Text = GridViewToolManagement.GetFocusedRowCellValue("DELETE_REASON") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("DELETE_REASON").ToString();
            TxtManager1.Text = GridViewToolManagement.GetFocusedRowCellValue("MANAGER1") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("MANAGER1").ToString();
            TxtManager2.Text = GridViewToolManagement.GetFocusedRowCellValue("MANAGER2") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("MANAGER2").ToString();
            TxtUserCompo.Text = GridViewToolManagement.GetFocusedRowCellValue("USE_COMPO") == null ? "" : GridViewToolManagement.GetFocusedRowCellValue("USE_COMPO").ToString();



            PictureEdit pic = PictureEditTool;

            //string column = "TOOL_IMG";

            pic.Image = null;

            try
            {
                ImgData = (byte[])GridViewToolManagement.GetFocusedRowCellValue("TOOL_IMG") == null ? null : (byte[])GridViewToolManagement.GetFocusedRowCellValue("TOOL_IMG");
                if (ImgData != null) {
                    MemoryStream ms = new MemoryStream(ImgData);
                    Image returnImage = Image.FromStream(ms);
                    pic.Image = returnImage;
                    pic.Cursor = Cursors.Default;
                }
            }
            catch (InvalidCastException)
            {

            }

            // history 메서드
            GetToolHistory(sMg_No);
        }

        private void GetToolHistory(string sMg_No)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT A.MG_NO ");
            strSql.AppendLine("      , A.MG_HIS_SEQ");
            strSql.AppendLine("      , A.OCCUR_DT ");
            strSql.AppendLine("      , A.MG_DESC ");
            strSql.AppendLine("      , A.MG_COST ");
            strSql.AppendLine("      , A.MG_USER ");
            strSql.AppendLine("   FROM EQUIP_TOOL_MGT_HISTORY A");
            strSql.AppendLine("  WHERE A.MG_NO = '" + sMg_No + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridToolRecord.DataSource = dt;

            //if (dt.Rows.Count > 0)
            //{
            //    DateEditRecord.EditValue = GridViewToolRecord.GetFocusedRowCellValue("OCCUR_DT").ToString();
            //    TxtRecordNm.EditValue = GridViewToolRecord.GetFocusedRowCellValue("MG_USER").ToString();
            //    TxtRecordCost.EditValue = GridViewToolRecord.GetFocusedRowCellValue("MG_COST").ToString();
            //    TxtRecordHistory.EditValue = GridViewToolRecord.GetFocusedRowCellValue("MG_DESC").ToString();
            //}
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            TxtToolNm.Text = null;
            LkupEditDept.EditValue = "";
            TxtModelNm.Text = null;
            TxtMgNo.Text = null;
            TxtStandard.Text = null;
            TxtToolPrc.Text = null;
            TxtMadeCompany.Text = null;
            DateEditPurDate.EditValue = DateTime.Now.ToString("yyyy-MM-dd");
            RdgbDeleteYn.EditValue = "N";
            TxtDeleteReason.Text = null;
            TxtManager1.Text = null;
            TxtManager2.Text = null;
            TxtUserCompo.Text = null;
            PictureEditTool.EditValue = null;

            TxtToolNm.Focus();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            byte[] byteImg = null;
            if (PictureEditTool.Image == null)
            {

            }
            else
            {
                Image img = PictureEditTool.Image;
                MemoryStream ms = new MemoryStream();
                img.Save(ms, img.RawFormat);
                byteImg = ms.ToArray();
            }

            string sMgNo = TxtMgNo.Text;
            string sToolNm = TxtToolNm.Text;
            string sMgDept = LkupEditDept.Text;
            string sModelNm = TxtModelNm.Text;
            string sStandard = TxtStandard.Text;
            string sToolPrc = TxtToolPrc.Text;
            string sMadeCompany = TxtMadeCompany.Text;
            string sPurDate = DateEditPurDate.EditValue.ToString();
            string sDelete_Yn = RdgbDeleteYn.EditValue.ToString();
            string sDeleteReason = TxtDeleteReason.Text;
            string sUserCompo = TxtUserCompo.Text;
            string sTxtManager1 = TxtManager1.Text;
            string sTxtManager2 = TxtManager2.Text;

            if (string.IsNullOrEmpty(sMgNo))
            {
                XtraMessageBox.Show("관리번호를 입력해주세요.");
                TxtMgNo.Focus();
                return;
            }

            Cursor = Cursors.WaitCursor;

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            try
            {
                strSql.AppendLine("MERGE INTO EQUIP_TOOL_MGT AS a                        ");
                strSql.AppendLine("    USING(SELECT                                      ");
                strSql.AppendLine("            MG_NO = '" + sMgNo + "'                                ");
                strSql.AppendLine("            , TOOL_NM = '" + sToolNm + "'                          ");
                strSql.AppendLine("            , MG_DEPT = '" + sMgDept + "'         ");
                strSql.AppendLine("            , MODEL_NM = '" + sModelNm + "'          ");
                strSql.AppendLine("            , STANDARD = '" + sStandard + "'            ");
                strSql.AppendLine("            , TOOL_PRC = '" + sToolPrc + "'            ");
                strSql.AppendLine("            , MADE_COMPANY = '" + sMadeCompany + "'        ");
                strSql.AppendLine("            , PUR_DATE = '" + sPurDate + "'  ");
                strSql.AppendLine("            , DELETE_YN = '" + sDelete_Yn + "'          ");
                strSql.AppendLine("            , DELETE_REASON = '" + sDeleteReason + "'       ");
                strSql.AppendLine("            , USE_COMPO = '" + sUserCompo + "'           ");
                strSql.AppendLine("            , MANAGER1 = '" + sTxtManager1 + "'            ");
                strSql.AppendLine("            , MANAGER2 = '" + sTxtManager2 + " '           ");
                if (byteImg != null)
                    strSql.AppendLine("            , TOOL_IMG = @imgName");
                else
                    strSql.AppendLine("            , TOOL_IMG = null ");
                strSql.AppendLine(") AS b");
                strSql.AppendLine("    ON a.MG_NO = b.MG_NO                              ");
                strSql.AppendLine("    WHEN MATCHED THEN UPDATE SET                      ");
                strSql.AppendLine("        MG_NO = '" + sMgNo + "'                                    ");
                strSql.AppendLine("            , TOOL_NM = '" + sToolNm + "'                          ");
                strSql.AppendLine("            , MG_DEPT = '" + sMgDept + "'         ");
                strSql.AppendLine("            , MODEL_NM = '" + sModelNm + "'          ");
                strSql.AppendLine("            , STANDARD = '" + sStandard + "'            ");
                strSql.AppendLine("            , TOOL_PRC = '" + sToolPrc + "'            ");
                strSql.AppendLine("            , MADE_COMPANY = '" + sMadeCompany + "'        ");
                strSql.AppendLine("            , PUR_DATE = '" + sPurDate + "'  ");
                strSql.AppendLine("            , DELETE_YN = '" + sDelete_Yn + "'          ");
                strSql.AppendLine("            , DELETE_REASON = '" + sDeleteReason + "'       ");
                strSql.AppendLine("            , USE_COMPO = '" + sUserCompo + "'           ");
                strSql.AppendLine("            , MANAGER1 = '" + sTxtManager1 + "'            ");
                strSql.AppendLine("            , MANAGER2 = '" + sTxtManager2 + " '           ");
                if(byteImg != null)
                    strSql.AppendLine("            , TOOL_IMG = @imgName      ");
                else
                    strSql.AppendLine("            , TOOL_IMG = null");
                strSql.AppendLine("    WHEN NOT MATCHED THEN INSERT(                     ");
                strSql.AppendLine("             MG_NO                                         ");
                strSql.AppendLine("               , TOOL_NM                              ");
                strSql.AppendLine("               , MG_DEPT                              ");
                strSql.AppendLine("               , MODEL_NM                             ");
                strSql.AppendLine("               , STANDARD                             ");
                strSql.AppendLine("               , TOOL_PRC                             ");
                strSql.AppendLine("               , MADE_COMPANY                         ");
                strSql.AppendLine("               , PUR_DATE                             ");
                strSql.AppendLine("               , DELETE_YN                            ");
                strSql.AppendLine("               , DELETE_REASON                        ");
                strSql.AppendLine("               , USE_COMPO                            ");
                strSql.AppendLine("               , MANAGER1                             ");
                strSql.AppendLine("               , MANAGER2                             ");
                strSql.AppendLine("               , TOOL_IMG)                            ");
                strSql.AppendLine("    VALUES(                                           ");
                strSql.AppendLine("             '" + sMgNo + "'                                            ");
                strSql.AppendLine("               , '" + sToolNm + "'                                 ");
                strSql.AppendLine("               , '" + sMgDept + "'                               ");
                strSql.AppendLine("               , '" + sModelNm + "'                                 ");
                strSql.AppendLine("               , '" + sStandard + "'                                   ");
                strSql.AppendLine("               , '" + sToolPrc + "'                                   ");
                strSql.AppendLine("               , '" + sMadeCompany + "'                                   ");
                strSql.AppendLine("               , '" + sPurDate + "'                         ");
                strSql.AppendLine("               , '" + sDelete_Yn + "'                                  ");
                strSql.AppendLine("               , '" + sDeleteReason + "'                                   ");
                strSql.AppendLine("               , '" + sUserCompo + "'                                   ");
                strSql.AppendLine("               , '" + sTxtManager1 + "'                                   ");
                strSql.AppendLine("               , '" + sTxtManager2 + "'                                   ");
                if (byteImg != null)
                    strSql.AppendLine("               , @imgName                          ");
                else
                    strSql.AppendLine("            , null");
                strSql.AppendLine("); ");

                /*
                strSql.AppendLine(" INSERT INTO EQUIP_TOOL_MGT  ");
                strSql.AppendLine("           ( MG_NO ");
                strSql.AppendLine("           , TOOL_NM ");
                strSql.AppendLine("           , MG_DEPT");
                strSql.AppendLine("           , MODEL_NM");
                strSql.AppendLine("           , STANDARD");
                strSql.AppendLine("           , TOOL_PRC");
                strSql.AppendLine("           , MADE_COMPANY");
                strSql.AppendLine("           , PUR_DATE");
                strSql.AppendLine("           , DELETE_YN");
                strSql.AppendLine("           , DELETE_REASON");
                strSql.AppendLine("           , USE_COMPO");
                strSql.AppendLine("           , MANAGER1");
                strSql.AppendLine("           , MANAGER2");
                strSql.AppendLine("           , TOOL_IMG");
                strSql.AppendLine("           ) ");
                strSql.AppendLine(" VALUES    ( '" + sMgNo + "' ");
                strSql.AppendLine("           , '" + sToolNm + "' ");
                strSql.AppendLine("           , '" + sMgDept + "' ");
                strSql.AppendLine("           , '" + sModelNm + "' ");
                strSql.AppendLine("           , '" + sStandard + "' ");
                strSql.AppendLine("           , '" + sToolPrc + "' ");
                strSql.AppendLine("           , '" + sMadeCompany + "' ");
                strSql.AppendLine("           , '" + sPurDate + "' ");
                strSql.AppendLine("           , '" + sDelete_Yn + "' ");
                strSql.AppendLine("           , '" + sDeleteReason + "' ");
                strSql.AppendLine("           , '" + sUserCompo + "' ");
                strSql.AppendLine("           , '" + sTxtManager1 + "' ");
                strSql.AppendLine("           , '" + sTxtManager2 + "' ");
                strSql.AppendLine("           ,   @ImgName  ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine(" ON DUPLICATE KEY UPDATE TOOL_NM    = '" + sToolNm + "' ");
                strSql.AppendLine("                       , MG_DEPT     = '" + sMgDept + "' ");
                strSql.AppendLine("                       , MODEL_NM     = '" + sModelNm + "' ");
                strSql.AppendLine("                       , STANDARD     = '" + sStandard + "' ");
                strSql.AppendLine("                       , TOOL_PRC     = '" + sToolPrc + "' ");
                strSql.AppendLine("                       , MADE_COMPANY     = '" + sMadeCompany + "' ");
                strSql.AppendLine("                       , PUR_DATE     = '" + sPurDate + "' ");
                strSql.AppendLine("                       , DELETE_YN     = '" + sDelete_Yn + "' ");
                strSql.AppendLine("                       , DELETE_REASON     = '" + sDeleteReason + "' ");
                strSql.AppendLine("                       , USE_COMPO     = '" + sUserCompo + "' ");
                strSql.AppendLine("                       , MANAGER1     = '" + sTxtManager1 + "' ");
                strSql.AppendLine("                       , MANAGER2     = '" + sTxtManager2 + " '");
                strSql.AppendLine("                       , TOOL_IMG     = @imgName");
                */

                cmd.CommandType = CommandType.Text;
                if (byteImg != null)
                    cmd.Parameters.AddWithValue("@imgName", byteImg);
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);

                GridViewToolManagement.FocusedRowHandle = GridViewToolManagement.LocateByDisplayText(0, GridColToolNm, sToolNm);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }

            Cursor = Cursors.Default;

        }
        private void BtnImgUp_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string image_file = string.Empty;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = @"C:\";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                image_file = dialog.FileName;
                PictureEditTool.Image = Bitmap.FromFile(image_file);
            }
        }

        private void BtnRecordNew_Click(object sender, EventArgs e)
        {
            DateEditRecord.EditValue = DateTime.Now.ToString("yyyy-MM-dd");
            TxtRecordNm.Text = null;
            TxtRecordCost.Text = null;
            TxtRecordHistory.Text = null;

            DateEditRecord.ReadOnly = false;
            TxtRecordNm.ReadOnly = false;
            TxtRecordCost.ReadOnly = false;
            TxtRecordHistory.ReadOnly = false;

            BtnRecordCancel.Visible = true;
            layoutControlItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            layoutControlItem24.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            layoutControlItem22.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            TxtRecordNm.Focus();

        }

        private void BtnRecordSave_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();

            string sMgNo = TxtMgNo.Text;
            string sRecordDate = DateEditRecord.EditValue.ToString();
            string sRecordNm = TxtRecordNm.Text;
            string sRecordCost = TxtRecordCost.Text;
            string sRecordHistory = TxtRecordHistory.Text;
            string sMg_No = GridViewToolManagement.GetFocusedRowCellValue("MG_NO").ToString();

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;


            strSql.AppendLine(" INSERT INTO EQUIP_TOOL_MGT_HISTORY  ");
            strSql.AppendLine("           ( MG_NO ");
            strSql.AppendLine("           , OCCUR_DT");
            strSql.AppendLine("           , MG_DESC");
            strSql.AppendLine("           , MG_COST");
            strSql.AppendLine("           , MG_USER");
            strSql.AppendLine("           ) ");
            strSql.AppendLine(" VALUES    ( '" + sMgNo + "' ");
            strSql.AppendLine("           , '" + sRecordDate + "' ");
            strSql.AppendLine("           , '" + sRecordHistory + "' ");
            strSql.AppendLine("           , '" + sRecordCost + "' ");
            strSql.AppendLine("           , '" + sRecordNm + "' ");
            strSql.AppendLine("           ) ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            DBConn.dbTran.Commit();
            DBConn.dbTran = null;
            MessageBox.Show("저장을 완료했습니다.");

            BtnRecordCancel.PerformClick();
            GetToolHistory(sMg_No);
        }

        private void GridViewToolRecord_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            DateEditRecord.EditValue = GridViewToolRecord.GetFocusedRowCellValue("OCCUR_DT") is null ? "" : GridViewToolRecord.GetFocusedRowCellValue("OCCUR_DT").ToString();
            TxtRecordNm.Text = GridViewToolRecord.GetFocusedRowCellValue("MG_USER") is null ? "" : GridViewToolRecord.GetFocusedRowCellValue("MG_USER").ToString();
            TxtRecordCost.Text = GridViewToolRecord.GetFocusedRowCellValue("MG_COST") is null ? "" : GridViewToolRecord.GetFocusedRowCellValue("MG_COST").ToString();
            TxtRecordHistory.Text = GridViewToolRecord.GetFocusedRowCellValue("MG_DESC") is null ? "" : GridViewToolRecord.GetFocusedRowCellValue("MG_DESC").ToString();
        }

        private void BtnRecordCancel_Click(object sender, EventArgs e)
        {
            BtnRecordCancel.Visible = false;
            layoutControlItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            layoutControlItem24.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            layoutControlItem22.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

            DateEditRecord.EditValue = null;
            TxtRecordNm.Text = null;
            TxtRecordCost.Text = null;
            TxtRecordHistory.Text = null;

            DateEditRecord.ReadOnly = true;
            TxtRecordNm.ReadOnly = true;
            TxtRecordCost.ReadOnly = true;
            TxtRecordHistory.ReadOnly = true;
        }

        private void EquipToolMgt_KeyDown(object sender, KeyEventArgs e)
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
            else if(e.KeyCode == Keys.F4)
            {
                BtnDel.PerformClick();
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F7)
            {
                BtnRecordCancel_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
            else if (e.KeyCode == Keys.F9)
            {
                BtnRecordNew_Click(null, null);
            }
            else if (e.KeyCode == Keys.F11)
            {
                BtnRecordSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F12)
            {
                //BtnDtlExcel_Click(null, null);
            }
        }

        private void GridToolManagement_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewToolManagement_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
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
                string sFileNM = "치공구리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridToolManagement.ExportToXls(FileName + ".xls");
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
                string sFileNM = "치공구 수리개조 내역";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridToolRecord.ExportToXls(FileName + ".xls");
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

        private void RdgbMenuDeleteYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = GetLookUpData();
            ComGrid.SetLookUpEdit(LkupToolNmRetr, dt, "CD", "CD", "Y");
            LkupToolNmRetr.Properties.PopulateColumns();
            LkupToolNmRetr.Properties.Columns[1].Visible = false;
        }

        private void EquipToolMgt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void LkupToolNmRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            try
            {
                string sMG_NO = GridViewToolManagement.GetFocusedRowCellValue("MG_NO")?.ToString();

                if (string.IsNullOrEmpty(sMG_NO))
                    return;

                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM EQUIP_TOOL_MGT_HISTORY");
                strSql.AppendLine("  WHERE MG_NO = '"+ sMG_NO + "'                 ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM EQUIP_TOOL_MGT");
                strSql.AppendLine("  WHERE MG_NO = '"+ sMG_NO + "'         ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewToolManagement.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewToolManagement.FocusedRowHandle = idx - 1;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
    }
}