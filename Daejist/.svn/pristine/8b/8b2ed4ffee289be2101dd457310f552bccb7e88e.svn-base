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
using DevExpress.Data;
using DevExpress.XtraGrid.Views.Grid;
using ComLib;
using DevExpress.XtraGrid;
using MySql.Data.MySqlClient;
using DevExpress.DataAccess.Excel;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.EditForm.Helpers.Controls;
using System.Diagnostics;
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
*            
* 수정일자 : 2023-02-07
* 수정자   : 정은영
* 수정내용 : (현업요청)
*            1. 직송 수정, 삭제 하는 프로그램으로 변경
*            
*/
namespace AccAdm
{
    public partial class AccInOutMgt : DevExpress.XtraEditors.XtraForm
    {
        public AccInOutMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccInOutMgt_Load(object sender, EventArgs e)
        {
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccInOutMgt");
            

            LkupDealGb.EditValueChanged -= LkupDealGb_EditValueChanged;

            DtpRetrStrt.EditValue = DateTime.Now;
            DtpRetrEnd.EditValue = DateTime.Now;

            DataTable dtKera = GetLookUpData("1", "Y", "");
            DataTable dtKname = GetLookUpData("2", "Y", "");
            DataTable dtDealer = GetLookUpData("3", "Y", "");
            DataTable dtGubun1 = GetLookUpData("4", "Y", "");

            DataTable dtKeraView = GetLookUpData("1", "N", "");
            DataTable dtKnameView = GetLookUpData("2", "N", "");

            ComGrid.SetLookUpEdit(LkupDealGb, dtKera, "CD", "NM", "Y");
            ComGrid.SetLookUpEdit(LkupGbCd, dtKname, "CD", "NM", "Y");

            RepositoryItemGridLookUpEdit keraLkup = new RepositoryItemGridLookUpEdit();
            RepositoryItemGridLookUpEdit KnameLkup = new RepositoryItemGridLookUpEdit();
            RepositoryItemGridLookUpEdit dealerLkup = new RepositoryItemGridLookUpEdit();

            dealerLkup.Appearance.Font = new Font("맑은 고딕", 9);

            ComGrid.SetGridLookUpEdit(keraLkup, dtKeraView, GridRetr, GridColGb, "CD", "NM", "");
            ComGrid.SetGridLookUpEdit(KnameLkup, dtKnameView, GridRetr, GridColGea, "CD", "NM", "");
            ComGrid.SetGridLookUpEdit(dealerLkup, dtDealer, GridRetr, GridColInDealer, "CD", "NM", "");
            ComGrid.SetGridLookUpEdit(dealerLkup, dtDealer, GridRetr, GridColOutDealer, "CD", "NM", "");
            ComGrid.SetGridLookUpEdit(RepoGridLkupGubun1, dtGubun1, GridRetr, GridColGrade, "CD", "NM", "");
            

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

            LkupDealGb.EditValue = "직송";
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

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_NM AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'KERATYPE'");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_NM AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'K_NAME'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT CAST(DEALER_CD AS VARCHAR) AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
            }else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT J_SERIAL AS CD                             ");
                strSql.AppendLine("      , GUBUN1 AS NM                               ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY J_SERIAL) AS SEQ");
                strSql.AppendLine("   FROM JAJAE                                      ");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        string _sFileName;

        private void BtnFile_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 EXCEL 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            FileDialog fileDlg = new OpenFileDialog();

            try
            {
                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    _sFileName = fileDlg.FileName;

                    DevExpress.DataAccess.Excel.ExcelDataSource excelDataSource = new DevExpress.DataAccess.Excel.ExcelDataSource
                    {
                        FileName = _sFileName
                    };

                    ExcelWorksheetSettings workSheetSettings = new ExcelWorksheetSettings("sheet1");
                    excelDataSource.SourceOptions = new ExcelSourceOptions(workSheetSettings)
                    {
                        SkipEmptyRows = true,
                        UseFirstRowAsHeader = true
                    };

                    excelDataSource.Fill();
                    GridFile.DataSource = excelDataSource;

                    UploadData(GridViewFile);
                }
                //fileDlg.Dispose();
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

        private void UploadData(DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            try
            {
                for (int i = 0; i < gridView.RowCount; i++)
                {
                    string keratype = Convert.ToString(gridView.GetRowCellValue(i, "keratype"));
                    string maipcher = Convert.ToString(gridView.GetRowCellValue(i, "Maipcher"));
                    string jCompany = Convert.ToString(gridView.GetRowCellValue(i, "J_Company"));
                    string jBnum = Convert.ToString(gridView.GetRowCellValue(i, "J_Bnum"));
                    string agreeDate = Convert.ToString(gridView.GetRowCellValue(i, "Agree_Date"));
                    string sjTime = string.Empty;
                    if(DateTime.TryParse(agreeDate, out DateTime dateResult))
                    {
                        sjTime = dateResult.ToString("yyyy-MM-dd HH:mm;ss");
                        agreeDate = agreeDate.Substring(0, 10);
                    }
                    string gubun1 = Convert.ToString(gridView.GetRowCellValue(i, "Gubun1"));
                    double weight = Convert.ToDouble(gridView.GetRowCellValue(i, "weight"));
                    double idanga = Convert.ToDouble(gridView.GetRowCellValue(i, "idanga"));
                    double odanga = Convert.ToDouble(gridView.GetRowCellValue(i, "ODanga"));
                    double ikongkep = Convert.ToDouble(gridView.GetRowCellValue(i, "ikongkep"));
                    double okongkep = Convert.ToDouble(gridView.GetRowCellValue(i, "OKongkep"));
                    string jState2 = Convert.ToString(gridView.GetRowCellValue(i, "J_State2"));
                    double secondWeight = Convert.ToDouble(gridView.GetRowCellValue(i, "SecondWeight"));
                    double firstWeight = Convert.ToDouble(gridView.GetRowCellValue(i, "FirstWeight"));
                    double iweight = Convert.ToDouble(gridView.GetRowCellValue(i, "iweight"));
                    double oweight = Convert.ToDouble(gridView.GetRowCellValue(i, "OWeight"));
                    double ichagam = Convert.ToDouble(gridView.GetRowCellValue(i, "ichagam"));
                    double transportKumak = Convert.ToDouble(gridView.GetRowCellValue(i, "TransportKumak"));
                    string transportDanga = gridView.GetRowCellValue(i, "TransportDanga")?.ToString();


                    StringBuilder strSql = new StringBuilder();

                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO MESURING ");
                    strSql.AppendLine("           ( J_CHECK,        KERATYPE,          MAIPCHERID,   MAIPCHER,   J_ASSIGNID,       J_COMPANY,    SUN,          J_DATE,       FIRSTTIME, SECONDTIME ");
                    strSql.AppendLine("           , FIRSTWEIGHT,    SECONDWEIGHT,      WEIGHT,       J_SERIAL,   GUBUN1,           J_BNUM,       K_NAME,       ICHAGAM,      OCHAGAM,   IGAMGA ");
                    strSql.AppendLine("           , OGAMGA,         IWEIGHT,           OWEIGHT,      IDANGA,     ODANGA,           IKONGKEP,     OKONGKEP,     FUSERCODE,    USERCODE,  FBUSEOCODE ");
                    strSql.AppendLine("           , BUSEOCODE,      P_ID,              J_GARAGE,     U_DATE,     J_ID,             KYERYANG12,   DRIVER_INOUT, AGREE_DATE,   J_STATE2,  TRANSPORTDANGA ");
                    strSql.AppendLine("           , TRANSPORTKUMAK, TRANSPORTC_SERIAL, CUSTOMWEIGHT, LOSSWEIGHT, TRANSPORTJUNGSAN, IPCHULGOJ_ID, MAGAM_FLAG,   WEIGHT_GUBUN, LENGTHSID, DAMAGE ");
                    strSql.AppendLine("           , GUMSU_SERIAL,   HALINYUL,          SURYANG, ETC_DEALER_CD1, ETC_COST1  ");
                    strSql.AppendLine("           ) ");
                    strSql.AppendLine(" SELECT '' AS J_CHECK ");
                    strSql.AppendLine("      , '" + keratype + "' AS KERATYPE ");
                    strSql.AppendLine("      , A.DEALER_CD AS MAIPCHERID ");
                    strSql.AppendLine("      , A.DEALER_NM AS MAIPCHER ");
                    strSql.AppendLine("      , B.DEALER_CD AS J_ASSIGNID ");
                    strSql.AppendLine("      , B.DEALER_NM AS J_COMPANY ");
                    strSql.AppendLine("      , (SELECT ISNULL(MAX(SUN), 0) + 1 FROM MESURING WHERE KERATYPE = '" + keratype + "' AND J_DATE = '" + agreeDate + "') AS SUN ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS J_DATE ");
                    strSql.AppendLine("      , '" + sjTime + "' AS FIRSTTIME ");
                    strSql.AppendLine("      , '" + sjTime + "' AS SECONDTIME ");
                    strSql.AppendLine("      , " + firstWeight + " AS FIRSTWEIGHT ");
                    strSql.AppendLine("      , " + secondWeight + " AS SECONDWEIGHT ");
                    strSql.AppendLine("      , " + weight + " AS WEIGHT ");
                    strSql.AppendLine("      , 400 AS J_SERIAL -- HARD CODING  ");
                    strSql.AppendLine("      , '" + gubun1 + "' AS GUBUN1 ");
                    strSql.AppendLine("      , '" + jBnum + "' AS J_BNUM ");
                    strSql.AppendLine("      , '상품' AS K_NAME  -- HARD CODING ");
                    strSql.AppendLine("      , " + ichagam + " AS ICHAGAM ");
                    strSql.AppendLine("      , 0 AS OCHAGAM -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS IGAMGA -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS OGAMGA -- HARD CODING ");
                    strSql.AppendLine("      , " + iweight + " AS IWEIGHT ");
                    strSql.AppendLine("      , " + oweight + " AS OWEIGHT ");
                    strSql.AppendLine("      , " + idanga + " AS IDANGA ");
                    strSql.AppendLine("      , " + odanga + " AS ODANGA ");
                    strSql.AppendLine("      , " + ikongkep + " AS IKONGKEP ");
                    strSql.AppendLine("      , " + okongkep + " AS OKONGKEP ");
                    strSql.AppendLine("      , 0 AS FUSERCODE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS USERCODE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS FBUSEOCODE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS BUSEOCODE -- HARD CODING ");
                    strSql.AppendLine("      , 400 AS P_ID -- HARD CODING ");
                    strSql.AppendLine("      , 4 AS J_GARAGE -- HARD CODING ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS U_DATE ");
                    strSql.AppendLine("      , 0 AS J_ID -- HARD CODING ");
                    strSql.AppendLine("      , 1 AS KYERYANG12 -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS DRIVER_INOUT -- HARD CODING ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS AGREE_DATE ");
                    strSql.AppendLine("      , '" + jState2 + "' AS J_STATE2 ");
                    strSql.AppendLine("      , 0 AS TRANSPORTDANGA  -- HARD CODING ");
                    strSql.AppendLine("      , " + transportKumak + " AS TRANSPORTKUMAK ");
                    strSql.AppendLine("      , 0 AS TRANSPORTC_SERIAL -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS CUSTOMWEIGHT  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS LOSSWEIGHT  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS TRANSPORTJUNGSAN  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS IPCHULGOJ_ID -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS MAGAM_FLAG -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS WEIGHT_GUBUN -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS LENGTHSID -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS DAMAGE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS GUMSU_SERIAL  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS HALINYUL -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS SURYANG -- HARD CODING ");
                    strSql.AppendLine("      , C.DEALER_CD AS ETC_DEALER_CD1");//#00007
                    strSql.AppendLine("      , " + transportKumak + " AS ETC_COST1");//#00007
                    strSql.AppendLine("   FROM ACC_DEALER_CD A ");
                    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B ");
                    strSql.AppendLine("     ON B.DEALER_NM = '" + jCompany + "' ");
                    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C");
                    strSql.AppendLine("     ON C.DEALER_NM = '" + transportDanga + "'");
                    strSql.AppendLine("  WHERE A.DEALER_NM = '" + maipcher + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                if(gridView.RowCount > 0)
                {
                    MessageBox.Show("정상적으로 저장을 완료했습니다.");
                    string agreeDate = Convert.ToString(gridView.GetRowCellValue(0, "Agree_Date"));

                    DtpRetrStrt.EditValue = agreeDate;
                    DtpRetrEnd.EditValue = agreeDate;

                    BtnRetr.PerformClick();
                }
                else
                {
                    MessageBox.Show("저장할 자료가 없습니다.");
                }
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DtpRetrStrt, DtpRetrEnd))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DtpRetrStrt.EditValue = DtpRetrEnd.EditValue;
                return;
            }

            string sYmd = DtpRetrStrt.EditValue.ToString().Substring(0,10);
            string sEndYmd = DtpRetrEnd.EditValue.ToString().Substring(0, 10);
            string sKeraType = LkupDealGb.Text;
            string sGubun = LkupGbCd.Text;
            string sCloGb = CboCloGb.EditValue.ToString();
            //string sKeraType = LkupDealGb.EditValue.ToString();
            //string sGubun = LkupGbCd.EditValue.ToString();

            //GridRetr.DataSource = null;

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" SELECT  IFNULL(A.J_Check, '0') AS CLO  ");
            //strSql.AppendLine("       , A.SUN AS SUN   ");
            //strSql.AppendLine("       , DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS WEIGHT_YMD ");
            ////strSql.AppendLine("       , (SELECT X.COM_CD FROM COM_BASE_CD X WHERE X.CD_GB = 'KERATYPE' AND X.COM_NM = A.KERATYPE) AS GB ");
            ////strSql.AppendLine("       , (SELECT X.COM_CD FROM COM_BASE_CD X WHERE X.CD_GB = 'K_NAME' AND X.COM_NM = A.K_NAME) AS GEA ");
            //strSql.AppendLine("       , A.KERATYPE AS GB ");
            //strSql.AppendLine("       , A.K_NAME AS GEA ");
            //strSql.AppendLine("       , CONVERT(A.MAIPCHERID, CHAR) AS IN_DEALER_CD ");
            //strSql.AppendLine("       , CONVERT(A.J_ASSIGNID, CHAR) AS OUT_DEALER_CD ");
            //strSql.AppendLine("       , A.J_BNUM AS CARNO   ");
            ////strSql.AppendLine("       , (SELECT X.COM_CD FROM COM_BASE_CD X WHERE X.CD_GB = '0001' AND X.COM_NM = A.K_NAME) AS GRADE ");
            //strSql.AppendLine("       , A.GUBUN1 AS GRADE ");
            //strSql.AppendLine("       , A.SECONDWEIGHT AS TOT ");
            //strSql.AppendLine("       , A.FIRSTWEIGHT AS EMPAMT ");
            //strSql.AppendLine("       , A.WEIGHT AS OWN ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS REDUCE ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '출고' THEN A.OWEIGHT ELSE A.IWEIGHT END AS AMOUNT ");
            //strSql.AppendLine("       , A.CUSTOMWEIGHT AS CUSTOM ");
            //strSql.AppendLine("       , A.LOSSWEIGHT AS LOSS ");
            //strSql.AppendLine("       , A.IDANGA ");
            //strSql.AppendLine("       , A.ODANGA ");
            //strSql.AppendLine("       , A.IKONGKEP ");
            //strSql.AppendLine("       , A.OKONGKEP ");
            //strSql.AppendLine("       , A.TRANSPORTKUMAK AS COST ");
            //strSql.AppendLine("       , B.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("       , A.J_STATE AS RED_RMK ");
            //strSql.AppendLine("       , A.GUMSUBIGO AS GUM_RMK ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H : %I') ELSE DATE_FORMAT(A.SECONDTIME, '%H : %I') END AS TIME1 ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H : %I') ELSE DATE_FORMAT(A.FIRSTTIME, '%H : %I') END AS TIME2 ");
            //strSql.AppendLine("       , DATE_FORMAT(A.AGREE_DATE, '%Y-%m-%d') AS CLO_YMD ");
            //strSql.AppendLine("       , A.JUNPYOID AS JUNPYOID  ");
            //strSql.AppendLine("    FROM MESURING A     ");
            //strSql.AppendLine("    LEFT OUTER JOIN HR_EMP_BASIS B  ");
            //strSql.AppendLine("      ON A.GUMSU_SERIAL = B.EMP_ID  ");
            //strSql.AppendLine("   WHERE A.J_DATE >= '" + sYmd + "' ");
            //strSql.AppendLine("     AND A.J_DATE <= '" + sEndYmd + "' ");
            //if(sCloGb.Equals("마감"))
            //    strSql.AppendLine("     AND A.J_Check = '1' ");
            //if(sCloGb.Equals("미마감"))
            //    strSql.AppendLine("     AND A.J_Check = '' ");
            //strSql.AppendLine("     AND A.KERATYPE LIKE CONCAT('%', '" + sKeraType + "', '%') ");
            //if(!string.IsNullOrEmpty(sGubun)) strSql.AppendLine("     AND A.K_NAME = '" + sGubun + "' ");
            //strSql.AppendLine("   ORDER BY A.J_DATE, A.JUNPYOID ");
            #endregion

            strSql.AppendLine("SELECT ISNULL(A.J_Check, '') AS CLO                                            ");
            strSql.AppendLine("     , A.SUN AS SUN                                                             ");
            strSql.AppendLine("     , A.J_DATE AS WEIGHT_YMD                                                   ");
            strSql.AppendLine("     , A.KERATYPE AS GB                                                         ");
            strSql.AppendLine("     , A.K_NAME AS GEA                                                          ");
            strSql.AppendLine("     , A.MAIPCHERID AS IN_DEALER_CD                                             ");
            strSql.AppendLine("     , D1.DEALER_NM AS IN_DEALER_NM                                             ");//#0001
            strSql.AppendLine("     , A.J_ASSIGNID AS OUT_DEALER_CD                                            ");
            strSql.AppendLine("     , D2.DEALER_NM AS OUT_DEALER_NM                                             ");//#0001
            strSql.AppendLine("     , A.J_BNUM AS CARNO                                                        ");
            strSql.AppendLine("     , C.DAEGUBUN                                                        ");
            strSql.AppendLine("     , A.J_SERIAL AS GRADE                                                        ");
            strSql.AppendLine("     , C.GUBUN1                                                         ");//#0001
            strSql.AppendLine("     , A.SECONDWEIGHT AS TOT                                                    ");
            strSql.AppendLine("     , A.FIRSTWEIGHT AS EMPAMT                                                  ");
            strSql.AppendLine("     , A.WEIGHT AS OWN                                                          ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS REDUCE");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '출고' THEN A.OWEIGHT ELSE A.IWEIGHT END AS AMOUNT");
            strSql.AppendLine("     , A.CUSTOMWEIGHT AS CUSTOM                                                 ");
            strSql.AppendLine("     , A.LOSSWEIGHT AS LOSS                                                     ");
            strSql.AppendLine("     , A.IDANGA                                                                 ");
            strSql.AppendLine("     , A.ODANGA                                                                 ");
            strSql.AppendLine("     , A.IKONGKEP                                                               ");
            strSql.AppendLine("     , A.OKONGKEP                                                               ");
            strSql.AppendLine("     , A.TRANSPORTKUMAK AS COST                                                 ");
            strSql.AppendLine("     , B.EMP_NM AS INSPECTOR                                                    ");
            strSql.AppendLine("     , A.J_STATE AS RED_RMK                                                     ");
            strSql.AppendLine("     , A.GUMSUBIGO AS GUM_RMK                                                   ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN A.FIRSTTIME ELSE A.SECONDTIME END AS TIME1");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN A.SECONDTIME ELSE A.FIRSTTIME END AS TIME2");
            strSql.AppendLine("     , A.AGREE_DATE AS CLO_YMD                        ");
            strSql.AppendLine("     , A.JUNPYOID AS JUNPYOID                         ");
            strSql.AppendLine("     , A.IPCHULGO_MAIPID   ");//#0001
            strSql.AppendLine("     , A.IPCHULGO_MACHULID ");//#0001
            strSql.AppendLine("  FROM MESURING A                                     ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS B                      ");
            strSql.AppendLine("    ON A.GUMSU_SERIAL = B.EMP_ID                      ");
            strSql.AppendLine("  LEFT JOIN JAJAE C           ");
            strSql.AppendLine("    ON A.J_SERIAL = C.J_SERIAL");
            //#0001
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD D1       ");
            strSql.AppendLine("    ON A.MAIPCHERID = D1.DEALER_CD ");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD D2       ");
            strSql.AppendLine("    ON A.J_ASSIGNID = D2.DEALER_CD ");
            strSql.AppendLine(" WHERE A.J_DATE BETWEEN '" + sYmd + "' AND '" + sEndYmd + "' ");
            if(sCloGb.Equals("마감"))
                strSql.AppendLine("     AND A.J_Check = '1' ");
            if (sCloGb.Equals("미마감"))
                strSql.AppendLine("     AND A.J_Check = '' ");
            strSql.AppendLine("     AND A.KERATYPE LIKE CONCAT('%', '" + sKeraType + "', '%') ");
            if (!string.IsNullOrEmpty(sGubun)) strSql.AppendLine("     AND A.K_NAME = '" + sGubun + "' ");
            strSql.AppendLine("   ORDER BY A.J_DATE, A.JUNPYOID ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            if (sKeraType.Equals("직송"))
            {
                GridViewRetr.OptionsBehavior.Editable = true;
            }
            else
            {
                GridViewRetr.OptionsBehavior.Editable = false;
            }
        }

        private void LkupDealGb_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void LkupGbCd_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        int unCheckedCount;

        private void GridViewRetr_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e)
        {
            GridView view = sender as GridView;

            int summaryID = Convert.ToInt32((e.Item as GridSummaryItem).Tag);

            if (e.SummaryProcess == CustomSummaryProcess.Start)
            {
                unCheckedCount = 0;
            }

            if (e.SummaryProcess == CustomSummaryProcess.Calculate)
            {
                if (e.FieldValue.Equals("0")) unCheckedCount++;
            }

            if (e.SummaryProcess == CustomSummaryProcess.Finalize)
            {
                e.TotalValue = unCheckedCount;
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if(GridViewRetr.RowCount > 0)
            {
                string sGb = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, GridColGb)?.ToString();
                if (string.IsNullOrEmpty(sGb))
                {
                    XtraMessageBox.Show("거래구분을 입력하세요.");
                    GridViewRetr.FocusedRowHandle = GridViewRetr.RowCount - 1;
                    GridViewRetr.FocusedColumn = GridColGb;
                    GridViewRetr.Focus();
                    return;
                }
            }

            GridViewRetr.FocusedRowChanged -= GridViewRetr_FocusedRowChanged;
            GridViewRetr.AddNewRow();
            GridViewRetr.SetFocusedRowCellValue(GridColYmd,DateTime.Today.ToString("yyyy-MM-dd"));
            GridViewRetr.SetFocusedRowCellValue(GridColGb, "직송");
            GridViewRetr.FocusedColumn = GridColYmd;
            GridViewRetr.OptionsBehavior.Editable = true;
            GridViewRetr.Focus();

        }

        //#0001
        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int[] selRows = GridViewRetr.GetSelectedRows();

            if (selRows.Length == 0)
            {
                XtraMessageBox.Show("삭제할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("정말로 선택한 데이터를 삭제하시겠습니까?"), "삭제확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            try
            {
                for (int i = 0; i < selRows.Length; i++)
                {
                    DataRow row = GridViewRetr.GetDataRow(selRows[i]);

                    string sjChk = GridViewRetr.GetRowCellValue(selRows[i], "CLO")?.ToString();
                    string stype = GridViewRetr.GetRowCellValue(selRows[i], "GB")?.ToString();
                    string sJunpyoid = GridViewRetr.GetRowCellValue(selRows[i], "JUNPYOID")?.ToString();

                    if (!stype.Equals("직송"))//직송 자료만 관리하도록 하기로함
                    {
                        XtraMessageBox.Show("직송이 아닌 자료는 조회만 가능합니다.");
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        return;
                    }

                    //if (sjChk.Equals("1") && !stype.Equals("직송"))//마감된 직송이 아닌자료
                    //{
                    //    XtraMessageBox.Show("직송이 아닌 마감된 자료는 마감취소 후 삭제가 가능합니다.");
                    //    DBConn.dbTran.Rollback();
                    //    DBConn.dbTran = null;
                    //    return;
                    //}

                    strSql.AppendLine(" DELETE FROM IPCHULGO                                                          ");
                    strSql.AppendLine("  WHERE J_ID = (SELECT IPCHULGO_MAIPID FROM MESURING WHERE JUNPYOID = '" + sJunpyoid + "')  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    strSql.AppendLine(" DELETE FROM IPCHULGO                                                          ");
                    strSql.AppendLine("  WHERE J_ID = (SELECT IPCHULGO_MACHULID FROM MESURING WHERE JUNPYOID = '" + sJunpyoid + "')");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    strSql.AppendLine(" DELETE FROM INLIST                                                            ");
                    strSql.AppendLine("  WHERE J_ID = (SELECT IPCHULGO_MAIPID FROM MESURING WHERE JUNPYOID = '" + sJunpyoid + "')  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    strSql.AppendLine(" DELETE FROM INLIST                                                            ");
                    strSql.AppendLine("  WHERE J_ID = (SELECT IPCHULGO_MACHULID FROM MESURING WHERE JUNPYOID = '" + sJunpyoid + "')");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    strSql.AppendLine(" ");
                    strSql.AppendLine(" DELETE FROM MESURING ");
                    strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    string sREF_RMK = string.Empty;
                    string sRmk = "["+ row["GB"]?.ToString() + " 데이터 삭제]";

                    string sSTD_COLS = string.Format("{0}/{1}/{2}/차량:{3}/{4}/매출중량:{5}", row["WEIGHT_YMD"]?.ToString().Substring(0, 10), row["GB"]?.ToString(), row["OUT_DEALER_NM"], row["CARNO"], row["GUBUN1"], row["AMOUNT"]);
                    sREF_RMK = string.Format("TABLE : MESURING/JUNPYOID:{0}", sJunpyoid);

                    if (stype.Equals("직송")) //직송은 회계전표삭제
                    {
                        strSql.Clear();
                        strSql.AppendLine(" DELETE                ");
                        strSql.AppendLine("   FROM ACTRAN         ");
                        strSql.AppendLine("  WHERE REF3 = @JUNPYOID"); //ACTRAN, AAUTO -> A02 경우 직송으로 REF3(MesuringJunpyo)로 찾아 삭제
                        strSql.AppendLine("    AND AAUTO = 'A02' ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoid);
                        cmd.ExecuteNonQuery();

                        sREF_RMK += string.Format(", IPCHULGO/J_ID:({0},{1}), INLIST/J_ID:({0},{1})", row["IPCHULGO_MAIPID"]?.ToString(), row["IPCHULGO_MACHULID"]?.ToString());
                        sREF_RMK += ", ACTRAN/REF3:{"+ sJunpyoid + "}";
                    }
                    else if (stype.Equals("입고"))
                    {
                        sREF_RMK += string.Format(", IPCHULGO/J_ID:{0}, INLIST/J_ID:{0}", row["IPCHULGO_MAIPID"]?.ToString());
                    }
                    if (stype.Equals("출고"))
                    {

                        sREF_RMK += string.Format(", IPCHULGO/J_ID:{0}, INLIST/J_ID:{0}", row["IPCHULGO_MACHULID"]?.ToString());
                    }

                    strSql.Clear();
                    strSql.AppendLine("   INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK, STD_COLS, REF_RMK )   ");
                    strSql.AppendLine(" 	      VALUES ");
                    strSql.AppendLine(" 	           ( CONVERT(VARCHAR(20),GETDATE(),20) ");
                    strSql.AppendLine(" 	           , '" + FmMainToolBar2.UserID + "' ");
                    strSql.AppendLine(" 	           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    strSql.AppendLine("                     WHERE X1.OCCUR_DATE = CONVERT(VARCHAR(10),GETDATE(),23) ");
                    strSql.AppendLine("                       AND X1.USRCD = '" + FmMainToolBar2.UserID + "' ) --LOG_SEQ(구분자) ");
                    strSql.AppendLine(" 	           , '" + this.Name + "' ");
                    strSql.AppendLine(" 	           , 'D' ");
                    strSql.AppendLine(" 	           , '" + ClsFunc.GetLocalIP() + "' ");
                    strSql.AppendLine(" 	           , '" + sRmk + "' ");
                    strSql.AppendLine(" 	           , '" + sSTD_COLS + "' ");
                    strSql.AppendLine(" 	           , '" + sREF_RMK + "' ); ");

                    cmd.Parameters.Clear();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("정상적으로 삭제했습니다.");

                BtnRetr_Click(null, null);
                GridViewRetr.FocusedRowHandle = selRows[0]-1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            DataTable dt = (DataTable)GridRetr.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            //string sId = "MIG";
            //string agreeDate = string.Empty;
            //string jState2 = string.Empty;

            if (dtMerge.Rows.Count > 0)  // modify
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    StringBuilder strSql = new StringBuilder();

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        string sJunpyoId = string.Empty;
                        string jdate = string.Empty;
                        string jtime = string.Empty;
                        string keratype = string.Empty;
                        string kname = string.Empty;
                        string maipcher = string.Empty;
                        string jCompany = string.Empty;
                        string jBnum = string.Empty;
                        string gubun1 = string.Empty;
                        double secondWeight = 0;
                        double firstWeight = 0;
                        double weight = 0;
                        double ichagam = 0;
                        double ochagam = 0;
                        double iweight = 0;
                        double oweight = 0;
                        double custom = 0;
                        double loss = 0;
                        double idanga = 0;
                        double odanga = 0;
                        double ikongkep = 0;
                        double okongkep = 0;
                        double transportKumak = 0;
                        string inspector = string.Empty;
                        string reducermk = string.Empty;
                        string gumrmk = string.Empty;

                        string sCLO = dtMerge.Rows[j]["CLO"]?.ToString();
                        sJunpyoId = dtMerge.Rows[j]["JUNPYOID"]?.ToString();
                        jdate = dtMerge.Rows[j]["WEIGHT_YMD"]?.ToString();
                        if(DateTime.TryParse(jdate, out DateTime dateResult))
                        {
                            jtime = dateResult.ToString("HH:mm");
                        }

                        jdate = jdate.Substring(0, 10);

                        keratype = dtMerge.Rows[j]["GB"]?.ToString();

                        if (!keratype.Equals("직송"))
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            XtraMessageBox.Show("직송이 아닌 자료는 수정 할 수 없습니다.");
                            return;
                        }

                        //if (!string.IsNullOrEmpty(sCLO) && sCLO.Equals("1") && !keratype.Equals("직송")) //직송이 아닌 마감된 자료는 저장x
                        //{
                        //    DBConn.dbTran.Rollback();
                        //    DBConn.dbTran = null;
                        //    XtraMessageBox.Show("직송이 아닌 마감된 자료는 수정 할 수 없습니다.");
                        //    return;
                        //}

                        kname = dtMerge.Rows[j]["GEA"]?.ToString();
                        maipcher = dtMerge.Rows[j]["IN_DEALER_CD"]?.ToString();
                        jCompany = dtMerge.Rows[j]["OUT_DEALER_CD"]?.ToString();
                        jBnum = dtMerge.Rows[j]["CARNO"]?.ToString();
                        gubun1 = dtMerge.Rows[j]["GRADE"]?.ToString();
                        double.TryParse(dtMerge.Rows[j]["TOT"]?.ToString(), out secondWeight);
                        double.TryParse(dtMerge.Rows[j]["EMPAMT"]?.ToString(), out firstWeight);
                        double.TryParse(dtMerge.Rows[j]["OWN"]?.ToString(), out weight);

                        if (keratype.Equals("출고"))
                        {
                            double.TryParse(dtMerge.Rows[j]["REDUCE"]?.ToString(), out ochagam);
                            double.TryParse(dtMerge.Rows[j]["AMOUNT"]?.ToString(), out oweight);
                            ichagam = 0;
                            iweight = 0;
                        }
                        else
                        {
                            ochagam = 0;
                            oweight = 0;
                            double.TryParse(dtMerge.Rows[j]["REDUCE"]?.ToString(), out ichagam);
                            double.TryParse(dtMerge.Rows[j]["AMOUNT"]?.ToString(), out iweight);
                        }

                        double.TryParse(dtMerge.Rows[j]["CUSTOM"]?.ToString(), out custom);
                        double.TryParse(dtMerge.Rows[j]["LOSS"]?.ToString(), out loss);
                        double.TryParse(dtMerge.Rows[j]["IDANGA"]?.ToString(), out idanga);
                        double.TryParse(dtMerge.Rows[j]["ODANGA"]?.ToString(), out odanga);
                        double.TryParse(dtMerge.Rows[j]["IKONGKEP"]?.ToString(), out ikongkep);
                        double.TryParse(dtMerge.Rows[j]["OKONGKEP"]?.ToString(), out okongkep);
                        double.TryParse(dtMerge.Rows[j]["COST"]?.ToString(), out transportKumak);
                        inspector = dtMerge.Rows[j]["INSPECTOR"]?.ToString();
                        reducermk = dtMerge.Rows[j]["RED_RMK"]?.ToString();
                        gumrmk = dtMerge.Rows[j]["GUM_RMK"]?.ToString();

                        string sSun = dtMerge.Rows[j]["SUN"]?.ToString();

                        if (string.IsNullOrEmpty(sSun))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT ISNULL(MAX(SUN) +1, 1) ");
                            strSql.AppendLine("   FROM MESURING ");
                            strSql.AppendLine("  WHERE J_DATE = CONVERT(DATE,'" + jdate + "') ");
                            strSql.AppendLine("    AND(('" + keratype + "' = '직송' AND KERATYPE = '직송') OR('" + keratype + "' <> '직송' AND KERATYPE <> '직송')) ");

                            DataTable dtSeq = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            sSun = dtSeq.Rows[0][0].ToString();
                        }

                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO MESURING ");
                        //strSql.AppendLine("           ( JUNPYOID,    J_CHECK,      KERATYPE,   MAIPCHERID, MAIPCHER, J_ASSIGNID, J_COMPANY, SUN,        J_DATE,         FIRSTTIME,  SECONDTIME ");
                        //strSql.AppendLine("           , FIRSTWEIGHT, SECONDWEIGHT, WEIGHT,     GUBUN1,     J_BNUM,   K_NAME,     ICHAGAM,   OCHAGAM,    IWEIGHT,        OWEIGHT");
                        //strSql.AppendLine("           , IDANGA,      ODANGA,       IKONGKEP,   OKONGKEP,   P_ID,     J_STATE,    J_GARAGE,  KYERYANG12, TRANSPORTKUMAK, CUSTOMWEIGHT ");
                        //strSql.AppendLine("           , LOSSWEIGHT,  MAGAM_FLAG,   GUMSUBIGO ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine(" VALUES    ( '" + sJunpyoId + "' ");
                        //strSql.AppendLine("           , '0' ");
                        //strSql.AppendLine("           , '" + keratype + "' ");
                        //strSql.AppendLine("           , '" + maipcher + "' ");
                        //strSql.AppendLine("           , (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE IDT_NO = '" + maipcher + "') ");
                        //strSql.AppendLine("           , '" + jCompany + "' ");
                        //strSql.AppendLine("           , (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE IDT_NO = '" + jCompany + "') ");
                        //strSql.AppendLine("           , '" + sSun + "' ");
                        ////strSql.AppendLine("           , (SELECT IFNULL(MAX(SUN), 1) ");
                        ////strSql.AppendLine("                FROM MESURING ");
                        ////strSql.AppendLine("               WHERE J_DATE = '" + jdate + "' ");
                        ////strSql.AppendLine("                 AND (('직송' = '" + keratype + "' AND KERATYPE = '" + keratype + "') OR ('직송' <> '" + keratype + "' AND KERATYPE <> '" + keratype + "'))) ");
                        //strSql.AppendLine("           , '" + jdate + "' ");
                        //strSql.AppendLine("           , (SELECT CONVERT('" + jdate + "', DATETIME)) ");
                        //strSql.AppendLine("           , (SELECT CONVERT('" + jdate + "', DATETIME)) ");
                        //strSql.AppendLine("           , " + firstWeight + " ");
                        //strSql.AppendLine("           , " + secondWeight + " ");
                        //strSql.AppendLine("           , " + weight + " ");
                        //strSql.AppendLine("           , '" + gubun1 + "' ");
                        //strSql.AppendLine("           , '" + jBnum + "' ");
                        //strSql.AppendLine("           , '" + kname + "' ");
                        //strSql.AppendLine("           , " + ichagam + " ");
                        //strSql.AppendLine("           , " + ochagam + " ");
                        //strSql.AppendLine("           , " + iweight + " ");
                        //strSql.AppendLine("           , " + oweight + " ");
                        //strSql.AppendLine("           , " + idanga + " ");
                        //strSql.AppendLine("           , " + odanga + " ");
                        //strSql.AppendLine("           , " + ikongkep + " ");
                        //strSql.AppendLine("           , " + okongkep + " ");
                        //strSql.AppendLine("           , '100' ");
                        //strSql.AppendLine("           , '" + reducermk + "' ");
                        //strSql.AppendLine("           , 1 ");
                        //strSql.AppendLine("           , 1 ");
                        //strSql.AppendLine("           , " + transportKumak + " ");
                        //strSql.AppendLine("           , " + custom + " ");
                        //strSql.AppendLine("           , " + loss + " ");
                        //strSql.AppendLine("           , '0' ");
                        //strSql.AppendLine("           , '" + gumrmk + "' ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine(" ON DUPLICATE KEY UPDATE KERATYPE = '" + keratype + "' ");
                        //strSql.AppendLine("           , MAIPCHERID = '" + maipcher + "' ");
                        //strSql.AppendLine("           , MAIPCHER = (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE IDT_NO = '" + maipcher + "') ");
                        //strSql.AppendLine("           , J_ASSIGNID = '" + jCompany + "' ");
                        //strSql.AppendLine("           , J_COMPANY = (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE IDT_NO = '" + jCompany + "') ");
                        //strSql.AppendLine("           , J_DATE = '" + jdate + "' ");
                        //strSql.AppendLine("           , FIRSTTIME = (SELECT CONVERT('" + jdate + "', DATETIME)) ");
                        //strSql.AppendLine("           , SECONDTIME = (SELECT CONVERT('" + jdate + "', DATETIME)) ");
                        //strSql.AppendLine("           , FIRSTWEIGHT = " + firstWeight + " ");
                        //strSql.AppendLine("           , SECONDWEIGHT = " + secondWeight + " ");
                        //strSql.AppendLine("           , WEIGHT = " + weight + " ");
                        //strSql.AppendLine("           , GUBUN1 = '" + gubun1 + "' ");
                        //strSql.AppendLine("           , J_BNUM = '" + jBnum + "' ");
                        //strSql.AppendLine("           , K_NAME = '" + kname + "' ");
                        //strSql.AppendLine("           , ICHAGAM = " + ichagam + " ");
                        //strSql.AppendLine("           , OCHAGAM = " + ochagam + " ");
                        //strSql.AppendLine("           , IWEIGHT = " + iweight + " ");
                        //strSql.AppendLine("           , OWEIGHT = " + oweight + " ");
                        //strSql.AppendLine("           , IDANGA = " + idanga + " ");
                        //strSql.AppendLine("           , ODANGA = " + odanga + " ");
                        //strSql.AppendLine("           , IKONGKEP = " + ikongkep + " ");
                        //strSql.AppendLine("           , OKONGKEP = " + okongkep + " ");
                        //strSql.AppendLine("           , P_ID = '100' ");
                        //strSql.AppendLine("           , J_STATE = '" + reducermk + "' ");
                        //strSql.AppendLine("           , J_GARAGE = 1 ");
                        //strSql.AppendLine("           , KYERYANG12 = 1 ");
                        //strSql.AppendLine("           , TRANSPORTKUMAK = " + transportKumak + " ");
                        //strSql.AppendLine("           , CUSTOMWEIGHT = " + custom + " ");
                        //strSql.AppendLine("           , LOSSWEIGHT = " + loss + " ");
                        //strSql.AppendLine("           , MAGAM_FLAG = '0' ");
                        //strSql.AppendLine("           , GUMSUBIGO = '" + gumrmk + "' ");
                        #endregion

                        strSql.Clear();
                        strSql.AppendLine("IF EXISTS(SELECT * FROM MESURING WHERE JUNPYOID = '"+ sJunpyoId + "') ");
                        strSql.AppendLine("   BEGIN                                            ");
                        strSql.AppendLine("       UPDATE MESURING                              ");
                        strSql.AppendLine("           SET KERATYPE = '" + keratype + "' ");
			            strSql.AppendLine("             , MAIPCHERID = '" + maipcher + "' ");
			            strSql.AppendLine("             , MAIPCHER = (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE CONVERT(VARCHAR, CONVERT(NUMERIC,DEALER_CD)) = '" + maipcher + "') ");
			            strSql.AppendLine("             , J_ASSIGNID = '" + jCompany + "' ");
			            strSql.AppendLine("             , J_COMPANY = (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE CONVERT(VARCHAR, CONVERT(NUMERIC,DEALER_CD)) = '" + jCompany + "') ");
			            strSql.AppendLine("             , J_DATE = '" + jdate + "' ");
			            strSql.AppendLine("             , FIRSTTIME = '" + jtime + "' ");
			            strSql.AppendLine("             , SECONDTIME = '" + jtime + "' ");
			            strSql.AppendLine("             , FIRSTWEIGHT = " + firstWeight + " ");
			            strSql.AppendLine("             , SECONDWEIGHT = " + secondWeight + " ");
			            strSql.AppendLine("             , WEIGHT = " + weight + " ");
                        strSql.AppendLine("             , J_SERIAL = '"+ gubun1 + "'");
			            strSql.AppendLine("             , GUBUN1 = (SELECT GUBUN1 FROM JAJAE WHERE J_SERIAL = '" + gubun1 + "') ");
			            strSql.AppendLine("             , J_BNUM = '" + jBnum + "' ");
			            strSql.AppendLine("             , K_NAME = '" + kname + "' ");
			            strSql.AppendLine("             , ICHAGAM = " + ichagam + " ");
			            strSql.AppendLine("             , OCHAGAM = " + ochagam + " ");
			            strSql.AppendLine("             , IWEIGHT = " + iweight + " ");
			            strSql.AppendLine("             , OWEIGHT = " + oweight + " ");
			            strSql.AppendLine("             , IDANGA = " + idanga + " ");
			            strSql.AppendLine("             , ODANGA = " + odanga + " ");
			            strSql.AppendLine("             , IKONGKEP = " + ikongkep + " ");
			            strSql.AppendLine("             , OKONGKEP = " + okongkep + " ");
			            strSql.AppendLine("             , P_ID = '100' ");
			            strSql.AppendLine("             , J_STATE = '" + reducermk + "' ");
			            strSql.AppendLine("             , J_GARAGE = 1 ");
			            strSql.AppendLine("             , KYERYANG12 = 1 ");
			            strSql.AppendLine("             , TRANSPORTKUMAK = " + transportKumak + " ");
			            strSql.AppendLine("             , CUSTOMWEIGHT = " + custom + " ");
			            strSql.AppendLine("             , LOSSWEIGHT = " + loss + " ");
			            strSql.AppendLine("             , MAGAM_FLAG = '0' ");
			            strSql.AppendLine("             , GUMSUBIGO = '" + gumrmk + "' ");
                        strSql.AppendLine("         WHERE JUNPYOID = '"+ sJunpyoId + "' ");
                        strSql.AppendLine("   END                        ");
                        strSql.AppendLine("ELSE                          ");
                        strSql.AppendLine("    BEGIN                     ");
                        strSql.AppendLine("        INSERT INTO MESURING ");
                        strSql.AppendLine("                ( J_CHECK, KERATYPE, MAIPCHERID, MAIPCHER, J_ASSIGNID, J_COMPANY, SUN, J_DATE, FIRSTTIME, SECONDTIME ");
                        strSql.AppendLine("                , FIRSTWEIGHT, SECONDWEIGHT, WEIGHT, J_SERIAL, GUBUN1, J_BNUM, K_NAME, ICHAGAM, OCHAGAM, IWEIGHT, OWEIGHT");
                        strSql.AppendLine("                , IDANGA, ODANGA, IKONGKEP, OKONGKEP, P_ID, J_STATE, J_GARAGE, KYERYANG12, TRANSPORTKUMAK, CUSTOMWEIGHT ");
                        strSql.AppendLine("                , LOSSWEIGHT, MAGAM_FLAG, GUMSUBIGO ");
                        strSql.AppendLine("                ) ");
                        strSql.AppendLine("    VALUES(    '' ");
                        strSql.AppendLine("                , '" + keratype + "' ");
                        strSql.AppendLine("                , '" + maipcher + "' ");
                        strSql.AppendLine("                , (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE CONVERT(VARCHAR, CONVERT(NUMERIC,DEALER_CD)) = '" + maipcher + "') ");
				        strSql.AppendLine("                , '" + jCompany + "' ");
				        strSql.AppendLine("                , (SELECT DEALER_NM FROM ACC_DEALER_CD WHERE CONVERT(VARCHAR, CONVERT(NUMERIC,DEALER_CD)) = '" + jCompany + "') ");
				        strSql.AppendLine("                , '" + sSun + "' ");
				        strSql.AppendLine("                , '" + jdate + "' ");
				        strSql.AppendLine("                , '" + jtime + "' ");
				        strSql.AppendLine("                , '" + jtime + "' ");
				        strSql.AppendLine("                , " + firstWeight + " ");
				        strSql.AppendLine("                , " + secondWeight + " ");
				        strSql.AppendLine("                , " + weight + " ");
				        strSql.AppendLine("                , '" + gubun1 + "' ");
                        strSql.AppendLine("                , (SELECT GUBUN1 FROM JAJAE WHERE J_SERIAL = '" + gubun1 + "') ");
				        strSql.AppendLine("                , '" + jBnum + "' ");
				        strSql.AppendLine("                , '" + kname + "' ");
				        strSql.AppendLine("                , " + ichagam + " ");
				        strSql.AppendLine("                , " + ochagam + " ");
				        strSql.AppendLine("                , " + iweight + " ");
				        strSql.AppendLine("                , " + oweight + " ");
				        strSql.AppendLine("                , " + idanga + " ");
				        strSql.AppendLine("                , " + odanga + " ");
				        strSql.AppendLine("                , " + ikongkep + " ");
				        strSql.AppendLine("                , " + okongkep + " ");
				        strSql.AppendLine("                , '100' ");
				        strSql.AppendLine("                , '" + reducermk + "' ");
				        strSql.AppendLine("                , 1 ");
				        strSql.AppendLine("                , 1 ");
				        strSql.AppendLine("                , " + transportKumak + " ");
				        strSql.AppendLine("                , " + custom + " ");
				        strSql.AppendLine("                , " + loss + " ");
				        strSql.AppendLine("                , '0' ");
				        strSql.AppendLine("                , '" + gumrmk + "' ");
				        strSql.AppendLine("                ) ");
                        strSql.AppendLine("    END");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    BtnRetr.PerformClick();
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    GridViewRetr.FocusedRowChanged += GridViewRetr_FocusedRowChanged;
                }
            }
            else
            {
                XtraMessageBox.Show("변경된 내역이 없습니다.");
            }
        }

        private void GridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.Equals("WEIGHT_YMD"))
            {
                string sYmd = e.Value.ToString().Replace("-", "").Substring(0, 8);
                e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
            }
        }

        private void repositoryItemDateEdit1_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            string sYmd = e.Value.ToString();

            if (sYmd.Length > 0)
            {
                sYmd = sYmd.Replace("-", "").Substring(0, 8);
            }
            else
            {
                sYmd = DateTime.Now.ToString("yyyyMMdd");
            }

            e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
        }

        private void GridViewRetr_EditFormShowing(object sender, EditFormShowingEventArgs e)
        {
            GridView view = sender as GridView;
            if (view == null) return;


            //if (view.GetFocusedRowCellValue("CLO").ToString().Equals("1"))
            //{
            //    MessageBox.Show("마감된 데이터는 수정할 수 없습니다.");
            //    e.Allow = false;
            //    return;
            //}

            if (!(view.GetFocusedRowCellDisplayText("GB").ToString().Equals("직송") || view.GetFocusedRowCellDisplayText("GB").ToString().Equals("")))
            {
                MessageBox.Show("거래구분이 '직송'이 아닌 데이터는 수정할 수 없습니다.");
                e.Allow = false;
                return;
            }
        }

        private void GridViewRetr_ShowingPopupEditForm(object sender, ShowingPopupEditFormEventArgs e)
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

        //private void EditFormUpdateButton_Click(object sender, EventArgs e)
        //{
        //    string sJunpyoid = GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "JUNPYOID").ToString();
        //    string sKeraType = GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "GB").ToString();
        //    double dIDanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "IDANGA"));
        //    double dODanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "ODANGA"));
        //    double dIKongKep = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "IKONGKEP"));
        //    double dOKongKep = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "OKONGKEP"));
        //    // double dAmount = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "AMOUNT"));

        //    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
        //    SqlCommand cmd = DBConn.dbCon.CreateCommand();
        //    cmd.Transaction = DBConn.dbTran;

        //    StringBuilder strSql = new StringBuilder();

        //    try
        //    {
        //        strSql = new StringBuilder();

        //        strSql.AppendLine(" UPDATE MESURING ");
        //        strSql.AppendLine("    SET IDANGA = " + dIDanGa);
        //        strSql.AppendLine("      , ODANGA = " + dODanGa);
        //        strSql.AppendLine("      , IKONGKEP = " + dIKongKep);
        //        strSql.AppendLine("      , OKONGKEP = " + dOKongKep);
        //        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'");

        //        cmd.CommandType = CommandType.Text;
        //        cmd.CommandText = strSql.ToString();
        //        cmd.ExecuteNonQuery();

        //        DBConn.dbTran.Commit();
        //        DBConn.dbTran = null;
        //        MessageBox.Show("저장을 완료했습니다.");
        //    }
        //    catch (Exception ex)
        //    {
        //        DBConn.dbTran.Rollback();
        //        DBConn.dbTran = null;
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName.Equals("IDANGA") || e.Column.FieldName.Equals("ODANGA") || e.Column.FieldName.Equals("REDUCE") || e.Column.FieldName.Equals("LOSS"))
            {
                string sKeraType = GridViewRetr.GetRowCellValue(e.RowHandle, "GB").ToString();
                double dWeight = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "OWN"));
                double dReduce = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "REDUCE"));
                double dLoss = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "LOSS"));
                double dIDanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "IDANGA"));
                double dODanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "ODANGA"));
                double dAmount = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "AMOUNT"));
                double dIKongKep = 0;
                double dOKongKep = 0;

                if (!sKeraType.Equals("출고"))  // 매입금액 수정
                {
                    dAmount = dWeight - dReduce - dLoss;
                    dIKongKep = dIDanGa * dAmount;
                }

                if (!sKeraType.Equals("입고")) // 매출금액 수정
                {
                    dAmount = dWeight - dReduce - dLoss;
                    dOKongKep = dODanGa * dAmount;
                }

                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "AMOUNT", dAmount);
                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "IKONGKEP", dIKongKep);
                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "OKONGKEP", dOKongKep);
            }

            GridViewRetr.UpdateCurrentRow();
        }

        private void SetColReadOnly(string sGb)
        {
            bool bReadOnly = sGb.Equals("Y") ? true : false;

            for(int i=0; i < GridViewRetr.Columns.Count; i++)
            {
                GridViewRetr.Columns[i].OptionsColumn.ReadOnly = bReadOnly;
            }

            GridViewRetr.Columns["CLO"].OptionsColumn.ReadOnly = true;
            GridViewRetr.Columns["SUN"].OptionsColumn.ReadOnly = true;
            GridViewRetr.Columns["JUNPYOID"].OptionsColumn.ReadOnly = true;
            GridViewRetr.Columns["AMOUNT"].OptionsColumn.ReadOnly = true;
            GridViewRetr.Columns["IKONGKEP"].OptionsColumn.ReadOnly = true;
            GridViewRetr.Columns["OKONGKEP"].OptionsColumn.ReadOnly = true;
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(GridViewRetr.RowCount > 0)
            {
                string sKeraType = GridViewRetr.GetRowCellValue(e.FocusedRowHandle, "GB")?.ToString();
                string sClo = GridViewRetr.GetRowCellValue(e.FocusedRowHandle, "CLO")?.ToString();

                if(!string.IsNullOrEmpty(sKeraType) && sKeraType.Equals("직송"))
                {
                    GridViewRetr.OptionsBehavior.Editable = true;
                }
                else
                {
                    GridViewRetr.OptionsBehavior.Editable = false;
                }

                //if (sClo.Equals("1"))
                //{
                //    //SetColReadOnly("Y"); //마감자료 수정불가처리
                //
                //    GridViewRetr.OptionsBehavior.Editable = false;
                //}
                //else
                //{
                //    GridViewRetr.OptionsBehavior.Editable = true; //미마감자료 수정 가능 처리
                //    //if (!sKeraType.Equals("직송"))
                //    //{
                //    //    SetColReadOnly("Y");
                //
                //    //    GridViewRetr.Columns["REDUCE"].OptionsColumn.ReadOnly = false;
                //    //    GridViewRetr.Columns["LOSS"].OptionsColumn.ReadOnly = false;
                //    //    GridViewRetr.Columns["COST"].OptionsColumn.ReadOnly = false;
                //    //    GridViewRetr.Columns["RED_RMK"].OptionsColumn.ReadOnly = false;
                //
                //    //    if (sKeraType.Equals("입고"))
                //    //    {
                //    //        GridViewRetr.Columns["IDANGA"].OptionsColumn.ReadOnly = false;
                //    //    }
                //    //    else
                //    //    {
                //    //        GridViewRetr.Columns["ODANGA"].OptionsColumn.ReadOnly = false;
                //    //    }
                //    //}
                //    //else
                //    //{
                //    //    SetColReadOnly("N");
                //
                //    //    GridViewRetr.Columns["CLO"].OptionsColumn.ReadOnly = true;
                //    //    GridViewRetr.Columns["SUN"].OptionsColumn.ReadOnly = true;
                //    //}
                //}
            }
        }

        private void GridViewRetr_FocusedColumnChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedColumnChangedEventArgs e)
        {
            if (GridViewRetr.RowCount > 0)
            {
                string sKeraType = GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "GB")?.ToString();

                if (!string.IsNullOrEmpty(sKeraType) && sKeraType.Equals("직송"))
                {
                    GridViewRetr.OptionsBehavior.Editable = true;
                }
                else
                {
                    GridViewRetr.OptionsBehavior.Editable = false;
                }
            }
        }

        private void AccInOutMgt_KeyDown(object sender, KeyEventArgs e)
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
                BtnDel_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnFile_Click(null, null); 
            }
            else if (e.KeyCode == Keys.F12)
            {
                BtnExcel_Click(null, null); 
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void AccInOutMgt_FormClosed(object sender, FormClosedEventArgs e)
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
                string sFileNM = "입출고 리스트";
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

        private void CboCloGb_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void AccInOutMgt_TextChanged(object sender, EventArgs e)
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

        private void CboCloGb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void RepoGridLkupGubun1_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEditBase lkupedit = (LookUpEditBase)sender;

            string sVal = lkupedit.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
                return;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT DAEGUBUN      ");
            strSql.AppendLine("   FROM JAJAE         ");
            strSql.AppendLine("  WHERE J_SERIAL = '"+ sVal + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt != null && dt.Rows.Count > 0)
            {
                string sDAEGUBUN = dt.Rows[0]["DAEGUBUN"]?.ToString();

                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, gridColumn1, sDAEGUBUN);
            }

            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridViewRetr.FocusedColumn, sVal);
        }

        private void repositoryItemDateEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            if(dateEdit != null)
            {
                string sVal = dateEdit.EditValue?.ToString();

                if (!string.IsNullOrEmpty(sVal))
                {
                    if(DateTime.TryParse(sVal, out DateTime result))
                    {
                        GridViewRetr.SetFocusedRowCellValue("WEIGHT_YMD", result.ToString("yyyy-MM-dd"));
                    }
                }
            }
        }
    }
}