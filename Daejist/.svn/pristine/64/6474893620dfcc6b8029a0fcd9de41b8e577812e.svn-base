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
using System.IO;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class CompanyInfo : DevExpress.XtraEditors.XtraForm
    {
        public CompanyInfo()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        private void CompanyInfo_Load(object sender, EventArgs e)
        {
            //
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccDealerCdDev");

            GetComnData();
        }

        private void GetComnData()
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append(" SELECT COMNM   ");
            strSql.Append("      , COMENM  ");
            strSql.Append("      , CVNAM   ");
            strSql.Append("      , SANO    ");
            strSql.Append("      , CVNO    ");
            strSql.Append("      , OWNAM   ");
            strSql.Append("      , UPTAE   ");
            strSql.Append("      , JONGK   ");
            strSql.Append("      , OPDAT   ");
            strSql.Append("      , ZIPCD   ");
            strSql.Append("      , ADDR1   ");
            strSql.Append("      , TELNO   ");
            strSql.Append("      , FAXNO   ");
            strSql.Append("      , EMAIL   ");
            strSql.Append("      , JONGCD   ");
            strSql.Append("      , TAXNM   ");
            strSql.Append("      , URL   ");
            strSql.Append("      , IMG   ");
            strSql.Append("   FROM COMPANYINFO ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtCompanyNm1.Text = dt.Rows[0]["COMNM"] is null ? "" : dt.Rows[0]["COMNM"].ToString();
            TxtCompanyEnglishNm.Text = dt.Rows[0]["COMENM"] is null ? "" : dt.Rows[0]["COMENM"].ToString();
            TxtCompanyNm2.Text = dt.Rows[0]["CVNAM"] is null ? "" : dt.Rows[0]["CVNAM"].ToString();
            TxtOwnerNm.Text = dt.Rows[0]["OWNAM"] is null ? "" : dt.Rows[0]["OWNAM"].ToString();
            TxtCompanyNo1.Text = dt.Rows[0]["SANO"] is null ? "" : dt.Rows[0]["SANO"].ToString();
            TxtCompanyNo2.Text = dt.Rows[0]["CVNO"] is null ? "" : dt.Rows[0]["CVNO"].ToString();
            TxtUPTAE.Text = dt.Rows[0]["UPTAE"] is null ? "" : dt.Rows[0]["UPTAE"].ToString();
            TxtJONGMOK.Text = dt.Rows[0]["JONGK"] is null ? "" : dt.Rows[0]["JONGK"].ToString();
            TxtPostNo.Text = dt.Rows[0]["ZIPCD"] is null ? "" : dt.Rows[0]["ZIPCD"].ToString();
            TxtAdress1.Text = dt.Rows[0]["ADDR1"] is null ? "" : dt.Rows[0]["ADDR1"].ToString();
            TxtTell1.Text = dt.Rows[0]["TELNO"] is null ? "" : dt.Rows[0]["TELNO"].ToString();
            TxtFax1.Text = dt.Rows[0]["FAXNO"] is null ? "" : dt.Rows[0]["FAXNO"].ToString();
            TxtEmail.Text = dt.Rows[0]["EMAIL"] is null ? "" : dt.Rows[0]["EMAIL"].ToString();
            DateEditOpenDt.EditValue = dt.Rows[0]["OPDAT"] is null ? "" : dt.Rows[0]["OPDAT"].ToString();
            TxtJONGMOKNo.Text = dt.Rows[0]["JONGCD"] is null ? "" : dt.Rows[0]["JONGCD"].ToString();
            TxtTaxNm.Text = dt.Rows[0]["TAXNM"] is null ? "" : dt.Rows[0]["TAXNM"].ToString();
            TxtUrl.Text = dt.Rows[0]["URL"] is null ? "" : dt.Rows[0]["URL"].ToString();
            PictureEditImg.Image = null;
            try
            {
                byte[] ImgData = (byte[])dt.Rows[0]["IMG"];
                MemoryStream ms = new MemoryStream(ImgData);
                Image returnImage = Image.FromStream(ms);
                PictureEditImg.Image = returnImage;
            }
            catch
            {

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
            else
            {
                return;
            }

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sCompanyNm1 = TxtCompanyNm1.Text.Equals("") ? "" : TxtCompanyNm1.Text;
            string sCompanyEnglishNm = TxtCompanyEnglishNm.Text.Equals("") ? "" : TxtCompanyEnglishNm.Text;
            string sCompanyNm2 = TxtCompanyNm2.Text.Equals("") ? "" : TxtCompanyNm2.Text;
            string sOwnerNm = TxtOwnerNm.Text.Equals("") ? "" : TxtOwnerNm.Text;
            string sCompanyNo1 = TxtCompanyNo1.Text.Equals("") ? "" : TxtCompanyNo1.Text;
            string sCompanyNo2 = TxtCompanyNo2.Text.Equals("") ? "" : TxtCompanyNo2.Text;
            string sUPTAE = TxtUPTAE.Text.Equals("") ? "" : TxtUPTAE.Text;
            string sJONGMOK = TxtJONGMOK.Text.Equals("") ? "" : TxtJONGMOK.Text;
            string sPostNo = TxtPostNo.Text.Equals("") ? "" : TxtPostNo.Text;
            string sAdress1 = TxtAdress1.Text.Equals("") ? "" : TxtAdress1.Text;
            string sTell1 = TxtTell1.Text.Equals("") ? "" : TxtTell1.Text;
            string sFax1 = TxtFax1.Text.Equals("") ? "" : TxtFax1.Text;
            string sEmail = TxtEmail.Text.Equals("") ? "" : TxtEmail.Text;
            string sOpenDt = DateEditOpenDt.EditValue.Equals("") ? "" : DateEditOpenDt.EditValue.ToString().Substring(0,10);
            string sJONGMOKNo = TxtJONGMOKNo.Text.Equals("") ? "" : TxtJONGMOKNo.Text;
            string sTaxNm = TxtTaxNm.Text.Equals("") ? "" : TxtTaxNm.Text;
            string sUrl = TxtUrl.Text.Equals("") ? "" : TxtUrl.Text;

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

                strSql.AppendLine("  UPDATE COMPANYINFO");
                strSql.AppendLine("     SET COMENM = '" + sCompanyNm1 + "'    ");
                strSql.AppendLine("        , CVNAM = '" + sCompanyNm2 + "'    ");
                strSql.AppendLine("        , SANO = '" + sCompanyNo1 + " '     ");
                strSql.AppendLine("        , CVNO = '" + sCompanyNo2 + " '     ");
                strSql.AppendLine("        , OWNAM = '" + sOwnerNm + " '    ");
                strSql.AppendLine("        , UPTAE = '" + sUPTAE + " '    ");
                strSql.AppendLine("        , JONGK = '" + sJONGMOK + " '    ");
                strSql.AppendLine("        , OPDAT = '" + sOpenDt + " '    ");
                strSql.AppendLine("        , ZIPCD = '" + sPostNo + " '    ");
                strSql.AppendLine("        , ADDR1 = '" + sAdress1 + " '    ");
                strSql.AppendLine("        , TELNO = '" + sTell1 + " '    ");
                strSql.AppendLine("        , FAXNO = '" + sFax1 + " '    ");
                strSql.AppendLine("        , EMAIL = '" + sEmail + "'    ");
                strSql.AppendLine("        , JONGCD = '" + sJONGMOKNo + "'   ");
                strSql.AppendLine("        , TAXNM = '" + sTaxNm + "'    ");
                strSql.AppendLine("        , URL = '" + sUrl + "'      ");
                strSql.AppendLine("        , IMG =@imgName        ");
                strSql.AppendLine("   WHERE COMNM = '" + sCompanyNm1 + "' ");


                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@imgName", byteImg);
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:COMPANYINFO -> COMENM:" + sCompanyNm1;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "U", this.Name , sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                GetComnData();

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
            Dispose();
        }

        private void CompanyInfo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                //BtnCrete_Click(null, null);
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
                //BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnPictureUp_Click(null, null);
            }
        }

        private void TxtUrl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSave.Focus();
        }
    }
}