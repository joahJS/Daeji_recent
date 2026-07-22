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
using DevExpress.XtraEditors.Repository;
using System.IO;
using MySql.Data.MySqlClient;
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
    public partial class EquipCdMgt : DevExpress.XtraEditors.XtraForm
    {
        public EquipCdMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void EquipCdMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DataTable dtDeptCd = GetLookUpData("1", "", "", "");
            ComLib.ComGrid.SetLookUpEdit(LkupMgDept, dtDeptCd, "CD", "NM", "Y");
            RepositoryItemGridLookUpEdit deptLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(deptLkup, dtDeptCd, GridRetr, GridColMgDept, "CD", "NM", "");

            DataTable dtEmp = GetLookUpData("2", "", "", "");
            ComLib.ComGrid.SetLookUpEdit(LkupMasterManager, dtEmp, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupSubManager, dtEmp, "CD", "NM", "Y");
            RepositoryItemGridLookUpEdit emptLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(emptLkup, dtEmp, GridRetr, GridColManagerMaster, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(emptLkup, dtEmp, GridRetr, GridColManagerSub, "CD", "NM", "");

            DataTable dtUserCd = GetLookUpData("3", "", "", "");
            RepositoryItemGridLookUpEdit userLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GridRetr, GridColEntID, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GridRetr, GridColMfyID, "CD", "NM", "");
            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "EquipCdMgt");

            DataTable dtEquipNm = GetLookUpData("4", "2", "", "");
            ComGrid.SetLookUpEdit(LkupRetrEquipNm,dtEquipNm,"CD","CD","Y");
            LkupRetrEquipNm.Properties.PopulateColumns();
            LkupRetrEquipNm.Properties.Columns[1].Visible = false;

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

            RdgbRetrUseYn.EditValue = "Y";
            BtnRetr_Click(null, null);
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT TOP 100 PERCENT A.DEPT_CD AS CD");
                //strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  ORDER BY SEQ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.EMP_ID AS CD");
                strSql.AppendLine("      , A.EMP_NM AS NM");
                strSql.AppendLine("      , A.EMP_ID AS SEQ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
                strSql.AppendLine("  WHERE 1=1");
                strSql.AppendLine("    AND A.EMPL_GB ='Y'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.USRCD AS CD");
                strSql.AppendLine("      , A.USRNM AS NM");
                strSql.AppendLine("      , A.USRCD AS SEQ");
                strSql.AppendLine("   FROM ZUSRLST A");
                strSql.AppendLine("  WHERE 1=1");
            }
            else if (sGb.Equals("4"))
            {
                string sUseYn = RdgbRetrUseYn.EditValue.ToString();

                strSql.AppendLine(" SELECT EQUIP_NM AS CD ");
                strSql.AppendLine(" , '' AS NM");
                strSql.AppendLine(" FROM EQUIP_CD   ");
                strSql.AppendLine("    WHERE USE_YN = '" + sUseYn + "'");
                strSql.AppendLine("GROUP BY EQUIP_NM");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");         

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }
       
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetGridRetr();
            GetGridRetrValue();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ClearAllForm();
            DateEditBuyDt.EditValue = DateTime.Now;
            RdgbUseYn.EditValue = "Y";
            TxtMgNo.Enabled = true;
            TxtMgNo.Focus();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sMgNo = TxtMgNo.EditValue?.ToString();
            if (string.IsNullOrEmpty(sMgNo))
            {
                XtraMessageBox.Show("관리번호를 입력하세요.");
                TxtMgNo.Focus();
                return;
            }
            SaveGridRetr();
        }

        private void TxtMgNo_Leave(object sender, EventArgs e)
        {
            if (Cursor.Current == BtnAdd.Cursor || Cursor.Current == BtnRetr.Cursor || Cursor.Current == BtnDelete.Cursor 
              || Cursor.Current == BtnClose.Cursor || Cursor.Current == BtnSave.Cursor)
            {
                return;
            }
            else
            {
                string sMgNo = TxtMgNo.EditValue?.ToString();
                if (string.IsNullOrEmpty(sMgNo))
                {
                    XtraMessageBox.Show("관리번호를 입력하세요.");
                    TxtMgNo.Focus();
                    TxtMgNo.SelectAll();
                    return;
                }

                if (CheckMgNoOverlappingValidation(sMgNo))
                {
                    XtraMessageBox.Show("해당 관리번호는 이미 존재하는 데이터입니다.\r\n다른 관리번호를 입력하세요.");
                    TxtMgNo.Focus();
                    TxtMgNo.SelectAll();
                    return;
                }
                return;
            }
        }

        private bool CheckMgNoOverlappingValidation(string MgNo)
        {
            bool bYN = false;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT COUNT(*) AS CNT_VALUE ");
            strSql.AppendLine("   FROM EQUIP_CD ");
            strSql.AppendLine("  WHERE MG_NO = '" + MgNo + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count > 0)
            {
                string sCntValue = dt.Rows[0]["CNT_VALUE"]?.ToString();
                if (string.IsNullOrEmpty(sCntValue))
                {
                    bYN = false;
                }
                else
                {
                    double dValue = Convert.ToDouble(sCntValue);
                    if(dValue == 0)
                    {
                        bYN = false;
                    }
                    else
                    {
                        bYN = true;
                    }
                }
            }
            else
            {
                bYN = false;
            }

            return bYN;
        }

        private void SaveGridRetr()
        {
            string sMgNo = TxtMgNo.EditValue?.ToString();
            string sEquipNm = TxtEquipNm.EditValue?.ToString();
            string sMgDept = LkupMgDept.EditValue?.ToString();
            string sModelNm = TxtModelNm.EditValue?.ToString();
            string sStandard = TxtStandard.EditValue?.ToString();
            string sBuyAmt = TxtBuyAmt.EditValue?.ToString();
            string sProducer = TxtProducer.EditValue?.ToString();
            string sBuyDt = DateEditBuyDt.EditValue?.ToString().Substring(0,10);
            string sUseYn = RdgbUseYn.EditValue?.ToString();
            string sDiscard = TxtDiscardMemo.EditValue?.ToString();
            string sUseCompo = TxtUseCompo.EditValue?.ToString();
            string sMasterManager = LkupMasterManager.EditValue?.ToString();
            string sMasterSub = LkupSubManager.EditValue?.ToString();
            string sUserCd = FmMainToolBar2.dtUserAutInfo.Rows[0]["USRCD"].ToString();

            if (string.IsNullOrEmpty(sMgNo))
            {
                XtraMessageBox.Show("관리번호를 입력하세요.");
                return;
            }
            else if (string.IsNullOrEmpty(sEquipNm))
            {
                XtraMessageBox.Show("설비명을 입력하세요.");
                return;
            }

            Image sImg = PictureEditImg.Image;

            byte[] byteImg = null;
            if (PictureEditImg.Image == null)
            {

            }
            else
            {
                Image img = PictureEditImg.Image;
                MemoryStream ms = new MemoryStream();
                img.Save(ms, img.RawFormat);
                byteImg = ms.ToArray();
            }

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            try
            {
                strSql.Clear();
                strSql.AppendLine("MERGE INTO EQUIP_CD AS a            ");
                strSql.AppendLine("    USING(SELECT                    ");
                strSql.AppendLine("        MG_NO = '" + sMgNo + "'               ");
                strSql.AppendLine("        , EQUIP_NM = '" + sEquipNm + "'          ");
                strSql.AppendLine("       , MG_DEPT = '" + sMgDept + "'           ");
                strSql.AppendLine("       , MODEL_NM = '" + sModelNm + "'            ");
                strSql.AppendLine("       , STANDARD = '" + sStandard + "'            ");
                strSql.AppendLine("       , EQUIP_PRC = '" + sBuyAmt + "'        ");
                strSql.AppendLine("       , MADE_COMPANY = '" + sProducer + "'        ");
                strSql.AppendLine("       , BUY_DATE = '"+ sBuyDt + "'    ");
                strSql.AppendLine("       , USE_YN = '" + sUseYn + "'               ");
                strSql.AppendLine("       , DELETE_REASON = '" + sDiscard + "'         ");
                strSql.AppendLine("       , USE_COMPO = '" + sUseCompo + "'             ");
                strSql.AppendLine("       , MANAGER_MASTER = '" + sMasterManager + "'        ");
                strSql.AppendLine("       , MANAGER_SUB = '" + sMasterSub + "'           ");
                strSql.AppendLine("       , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20)         ");
                strSql.AppendLine("       , MFY_ID = '" + sUserCd + "'               ");
                if (byteImg != null)
                    strSql.AppendLine("       , EQUIP_IMG = @imgName");
                strSql.AppendLine(") AS b ");
                strSql.AppendLine("    ON a.MG_NO = b.MG_NO            ");
                strSql.AppendLine("    WHEN MATCHED THEN UPDATE SET    ");
                strSql.AppendLine("        MG_NO = '" + sMgNo + "'               ");
                strSql.AppendLine("        , EQUIP_NM = '" + sEquipNm + "'          ");
                strSql.AppendLine("       , MG_DEPT = '" + sMgDept + "'           ");
                strSql.AppendLine("       , MODEL_NM = '" + sModelNm + "'            ");
                strSql.AppendLine("       , STANDARD = '" + sStandard + "'            ");
                strSql.AppendLine("       , EQUIP_PRC = '" + sBuyAmt + "'        ");
                strSql.AppendLine("       , MADE_COMPANY = '" + sProducer + "'        ");
                strSql.AppendLine("       , BUY_DATE = '" + sBuyDt + "'   ");
                strSql.AppendLine("       , USE_YN = '" + sUseYn + "'               ");
                strSql.AppendLine("       , DELETE_REASON = '" + sDiscard + "'         ");
                strSql.AppendLine("       , USE_COMPO = '" + sUseCompo + "'             ");
                strSql.AppendLine("       , MANAGER_MASTER = '" + sMasterManager + "'        ");
                strSql.AppendLine("       , MANAGER_SUB = '" + sMasterSub + "'           ");
                strSql.AppendLine("       , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20)          ");
                strSql.AppendLine("       , MFY_ID = '" + sUserCd + "'               ");
                if(byteImg != null)
                    strSql.AppendLine("       , EQUIP_IMG = @imgName       ");

                strSql.AppendLine("    WHEN NOT MATCHED THEN INSERT(   ");
                strSql.AppendLine("                MG_NO               ");
                strSql.AppendLine("               , EQUIP_NM           ");
                strSql.AppendLine("               , MG_DEPT            ");
                strSql.AppendLine("               , MODEL_NM           ");
                strSql.AppendLine("               , STANDARD           ");
                strSql.AppendLine("               , EQUIP_PRC          ");
                strSql.AppendLine("               , MADE_COMPANY       ");
                strSql.AppendLine("               , BUY_DATE           ");
                strSql.AppendLine("               , USE_YN             ");
                strSql.AppendLine("               , DELETE_REASON      ");
                strSql.AppendLine("               , USE_COMPO          ");
                strSql.AppendLine("               , MANAGER_MASTER     ");
                strSql.AppendLine("               , MANAGER_SUB        ");
                strSql.AppendLine("               , ENT_DT             ");
                strSql.AppendLine("               , ENT_ID             ");
                if(byteImg != null)
                    strSql.AppendLine("               , EQUIP_IMG         ");
                strSql.AppendLine(")");
                strSql.AppendLine("    VALUES(                         ");
                strSql.AppendLine("        '" + sMgNo + "'                       ");
                strSql.AppendLine("       , '" + sEquipNm + "'                      ");
                strSql.AppendLine("       , '" + sMgDept + "'                     ");
                strSql.AppendLine("       , '" + sModelNm + "'                       ");
                strSql.AppendLine("       , '" + sStandard + "'                       ");
                strSql.AppendLine("       , '" + sBuyAmt + "'                    ");
                strSql.AppendLine("       , '" + sProducer + "'                       ");
                strSql.AppendLine("       , '" + sBuyDt + "'                ");
                strSql.AppendLine("       , '" + sUseYn + "'                        ");
                strSql.AppendLine("       , '" + sDiscard + "'                         ");
                strSql.AppendLine("       , '" + sUseCompo + "'                         ");
                strSql.AppendLine("       , '" + sMasterManager + "'                         ");
                strSql.AppendLine("       , '" + sMasterSub + "'                         ");
                strSql.AppendLine("       , CONVERT(VARCHAR(20),GETDATE(),20)                 ");
                strSql.AppendLine("       , '" + sUserCd + "'                        ");
                if(byteImg != null)
                    strSql.AppendLine("       , @ImgName                 ");

                strSql.AppendLine("      );");


                /*
                strSql.AppendLine(" INSERT INTO EQUIP_CD ");
                strSql.AppendLine("           ( MG_NO ");
                strSql.AppendLine("           , EQUIP_NM ");
                strSql.AppendLine("           , MG_DEPT");
                strSql.AppendLine("           , MODEL_NM");
                strSql.AppendLine("           , STANDARD");
                strSql.AppendLine("           , EQUIP_PRC");
                strSql.AppendLine("           , MADE_COMPANY");
                strSql.AppendLine("           , BUY_DATE");
                strSql.AppendLine("           , USE_YN");
                strSql.AppendLine("           , DELETE_REASON");
                strSql.AppendLine("           , USE_COMPO");
                strSql.AppendLine("           , MANAGER_MASTER");
                strSql.AppendLine("           , MANAGER_SUB");
                strSql.AppendLine("           , ENT_DT");
                strSql.AppendLine("           , ENT_ID");
                strSql.AppendLine("           , MFY_DT");
                strSql.AppendLine("           , MFY_ID");
                strSql.AppendLine("           , EQUIP_IMG");
                strSql.AppendLine("           ) ");
                strSql.AppendLine(" VALUES    ( '" + sMgNo + "' ");
                strSql.AppendLine("           , '" + sEquipNm + "' ");
                strSql.AppendLine("           , '" + sMgDept + "' ");
                strSql.AppendLine("           , '" + sModelNm + "' ");
                strSql.AppendLine("           , '" + sStandard + "' ");
                strSql.AppendLine("           , '" + sBuyAmt + "' ");
                strSql.AppendLine("           , '" + sProducer + "' ");
                strSql.AppendLine("           , '" + sBuyDt + "' ");
                strSql.AppendLine("           , '" + sUseYn + "' ");
                strSql.AppendLine("           , '" + sDiscard + "' ");
                strSql.AppendLine("           , '" + sUseCompo + "' ");
                strSql.AppendLine("           , '" + sMasterManager + "' ");
                strSql.AppendLine("           , '" + sMasterSub + "' ");
                strSql.AppendLine("           ,   NOW()  ");
                strSql.AppendLine("           ,  '"+ sUserCd + "' ");
                strSql.AppendLine("           ,   NOW()  ");
                strSql.AppendLine("           ,   '" + sUserCd + "'  ");
                strSql.AppendLine("           ,   @ImgName  ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine(" ON DUPLICATE KEY UPDATE  ");
                strSql.AppendLine("                         EQUIP_NM    = '" + sEquipNm + "' ");
                strSql.AppendLine("                       , MG_DEPT     = '" + sMgDept + "' ");
                strSql.AppendLine("                       , MODEL_NM     = '" + sModelNm + "' ");
                strSql.AppendLine("                       , STANDARD     = '" + sStandard + "' ");
                strSql.AppendLine("                       , EQUIP_PRC     = '" + sBuyAmt + "' ");
                strSql.AppendLine("                       , MADE_COMPANY     = '" + sProducer + "' ");
                strSql.AppendLine("                       , BUY_DATE     = '" + sBuyDt + "' ");
                strSql.AppendLine("                       , USE_YN     = '" + sUseYn + "' ");
                strSql.AppendLine("                       , DELETE_REASON     = '" + sDiscard + "' ");
                strSql.AppendLine("                       , USE_COMPO     = '" + sUseCompo + "' ");
                strSql.AppendLine("                       , MANAGER_MASTER     = '" + sMasterManager+ "' ");
                strSql.AppendLine("                       , MANAGER_SUB     = '" +sMasterSub+ "'");
                strSql.AppendLine("                       , MFY_DT      = NOW()");
                strSql.AppendLine("                       , MFY_ID      =  '" + sUserCd + "' ");
                strSql.AppendLine("                       , EQUIP_IMG     = @imgName");
                */


                cmd.CommandType = CommandType.Text;
                if(byteImg != null)
                    cmd.Parameters.AddWithValue("@imgName", byteImg);

                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:EQUIP_CD -> MG_NO:" + sMgNo + "EQUIP_NM" + sEquipNm;
                ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, "1", "S", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColMgNo, sMgNo);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }

            Cursor = Cursors.Default;
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sMgNo = GridViewRetr.GetFocusedRowCellValue("MG_NO")?.ToString();
            string sMgNM = GridViewRetr.GetFocusedRowCellValue("EQUIP_NM")?.ToString();

            if (string.IsNullOrEmpty(sMgNo))
            {
                XtraMessageBox.Show("관리번호가 존재하지 않습니다.\r\n삭제하려는 데이터를 정확히 선택하세요.");
                return;
            }

            if (XtraMessageBox.Show("관리번호 : " + sMgNo + "\r\n설비명 : " + sMgNM + " \r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                , "설비 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }
            
            if(string.IsNullOrEmpty(sMgNo) || string.IsNullOrEmpty(sMgNM))
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
                strSql.AppendLine(" DELETE FROM EQUIP_CD ");
                strSql.AppendLine("       WHERE MG_NO = '" + sMgNo + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:EQUIP_CD -> MG_NO:" + sMgNo;
                ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr_Click(null, null);
                GridViewRetr.FocusedRowHandle = idx - 1;
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnPictureUp_Click(object sender, EventArgs e)
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
                PictureEditImg.Image = Bitmap.FromFile(image_file);
            }
        }

        private void GetGridRetr()
        {
            Cursor = Cursors.WaitCursor;
            StringBuilder strSql = new StringBuilder();
            string sEquipNm = LkupRetrEquipNm.Text;
            string sUseYn = RdgbRetrUseYn.EditValue.ToString();

            strSql.Clear();
            strSql.AppendLine("");
            strSql.AppendLine(" SELECT A.MG_NO  ");
            strSql.AppendLine(" 	 , A.EQUIP_NM ");
            strSql.AppendLine(" 	 , A.MG_DEPT");
            strSql.AppendLine(" 	 , A.MODEL_NM");
            strSql.AppendLine(" 	 , A.STANDARD");
            strSql.AppendLine(" 	 , A.EQUIP_PRC");
            strSql.AppendLine(" 	 , A.MADE_COMPANY");
            strSql.AppendLine(" 	 , A.BUY_DATE");
            strSql.AppendLine(" 	 , A.USE_YN");
            strSql.AppendLine(" 	 , A.DELETE_REASON");
            strSql.AppendLine(" 	 , A.USE_COMPO");
            strSql.AppendLine(" 	 , A.MANAGER_MASTER");
            strSql.AppendLine(" 	 , A.MANAGER_SUB");
            strSql.AppendLine(" 	 , A.EQUIP_IMG");
            strSql.AppendLine(" 	 , A.ENT_DT");
            strSql.AppendLine(" 	 , A.ENT_ID");
            strSql.AppendLine(" 	 , A.MFY_DT");
            strSql.AppendLine(" 	 , A.MFY_ID");          
            strSql.AppendLine("	  FROM EQUIP_CD A");
            strSql.AppendLine("  WHERE 1=1");
            if(sEquipNm != "")strSql.AppendLine("    AND EQUIP_NM = '" +  sEquipNm  + "'");
            strSql.AppendLine("    AND (('*' = '" + sUseYn + "') OR(('*' <> '" + sUseYn + "') AND USE_YN = '" + sUseYn + "'))");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridRetr.DataSource = dt;
            Cursor = Cursors.Default;
            
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sMgNo = GridViewRetr.GetFocusedRowCellValue("MG_NO")is null ? "" : GridViewRetr.GetFocusedRowCellValue("MG_NO").ToString();
            //string sMgNo =  GridViewRetr.GetFocusedRowCellValue("MG_NO").ToString();
            ClearAllForm();           
            GetGridRetrValue();
            TxtMgNo.Enabled = false;
        }

        private void GetGridRetrValue()
        {
            TxtMgNo.Text = GridViewRetr.GetFocusedRowCellValue("MG_NO") is null ? "" : GridViewRetr.GetFocusedRowCellValue("MG_NO").ToString();
            TxtEquipNm.Text = GridViewRetr.GetFocusedRowCellValue("EQUIP_NM") is null ? "" : GridViewRetr.GetFocusedRowCellValue("EQUIP_NM").ToString();
            LkupMgDept.EditValue = GridViewRetr.GetFocusedRowCellValue("MG_DEPT") is null ? "" : GridViewRetr.GetFocusedRowCellValue("MG_DEPT").ToString();
            TxtModelNm.Text = GridViewRetr.GetFocusedRowCellValue("MODEL_NM") is null ? "" : GridViewRetr.GetFocusedRowCellValue("MODEL_NM").ToString();
            TxtStandard.Text = GridViewRetr.GetFocusedRowCellValue("STANDARD") is null ? "" : GridViewRetr.GetFocusedRowCellValue("STANDARD").ToString();

            TxtBuyAmt.Text = GridViewRetr.GetFocusedRowCellValue("EQUIP_PRC") is null ? "" : GridViewRetr.GetFocusedRowCellValue("EQUIP_PRC").ToString();
            TxtProducer.Text = GridViewRetr.GetFocusedRowCellValue("MADE_COMPANY") is null ? "" : GridViewRetr.GetFocusedRowCellValue("MADE_COMPANY").ToString();
            DateEditBuyDt.EditValue = GridViewRetr.GetFocusedRowCellValue("BUY_DATE") is null ? "" : GridViewRetr.GetFocusedRowCellValue("BUY_DATE").ToString();
            RdgbUseYn.EditValue = GridViewRetr.GetFocusedRowCellValue("USE_YN") is null ? "" : GridViewRetr.GetFocusedRowCellValue("USE_YN").ToString();
            TxtDiscardMemo.Text = GridViewRetr.GetFocusedRowCellValue("DELETE_REASON") is null ? "" : GridViewRetr.GetFocusedRowCellValue("DELETE_REASON").ToString();
            TxtUseCompo.Text = GridViewRetr.GetFocusedRowCellValue("USE_COMPO") is null ? "" : GridViewRetr.GetFocusedRowCellValue("USE_COMPO").ToString();
            LkupMasterManager.EditValue = GridViewRetr.GetFocusedRowCellValue("MANAGER_MASTER") is null ? "" : GridViewRetr.GetFocusedRowCellValue("MANAGER_MASTER").ToString();
            LkupSubManager.EditValue = GridViewRetr.GetFocusedRowCellValue("MANAGER_SUB") is null ? "" : GridViewRetr.GetFocusedRowCellValue("MANAGER_SUB").ToString();
            PictureEditImg.Image = null;
            
            try
            {
                Byte[] ImgData = (byte[])GridViewRetr.GetFocusedRowCellValue("EQUIP_IMG");
                MemoryStream ms = new MemoryStream(ImgData);
                Image returnImage = Image.FromStream(ms);
                PictureEditImg.Image = returnImage;
            }
            catch
            {

            }            
        }
        public  void ClearAllForm()
        {            
            TxtMgNo.Text = null;
            TxtEquipNm.Text = null;
            LkupMgDept.EditValue = null;
            TxtModelNm.Text = null;
            TxtStandard.Text = null;
            TxtBuyAmt.Text = null;
            TxtProducer.Text = null;
            DateEditBuyDt.EditValue = null;
            RdgbUseYn.EditValue = null;
            TxtDiscardMemo.Text = null;
            TxtUseCompo.Text = null;
            LkupMasterManager.EditValue = null;
            LkupSubManager.EditValue = null;
            PictureEditImg.Image = null;
        }

        private void BtnHisSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            
        }
        
        private void EquipCdMgt_KeyDown(object sender, KeyEventArgs e)
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


        private void LkupRetrEquipNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridHistory_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RdgbRetrUseYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);

            DataTable dtEquipNm = GetLookUpData("4", "2", "", "");
            ComGrid.SetLookUpEdit(LkupRetrEquipNm, dtEquipNm, "CD", "CD", "Y");
            LkupRetrEquipNm.Properties.PopulateColumns();
            LkupRetrEquipNm.Properties.Columns[1].Visible = false;
        }

        private void EquipCdMgt_FormClosed(object sender, FormClosedEventArgs e)
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
                string sFileNM = "설비코드 리스트";
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

        private void EquipCdMgt_TextChanged(object sender, EventArgs e)
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
    }
}