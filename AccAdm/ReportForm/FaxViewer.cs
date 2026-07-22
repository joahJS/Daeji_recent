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
using Popbill.Fax;
using Popbill;
using DevExpress.XtraPrinting;
using System.IO;
using ComLib;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class FaxViewer : DevExpress.XtraEditors.XtraForm
    {
        private static DataTable dt;
        private string FaxName;
        private Dictionary<string, Image> dicImage;

        // 팩스제목: 팩스에 제목이 적히는건 아니고 팝빌 사이트에서 구분용
        public string faxTitle { get; set; }

        public FaxViewer()
        {
            InitializeComponent();
        }

        public FaxViewer(DataTable dt1, string Name)
        {
            InitializeComponent();
            dt = dt1;
            FaxName = Name;
        }

        public FaxViewer(DataTable dt1, string Name, Dictionary<string, Image> dicImage)
        {
            InitializeComponent();
            dt = dt1;
            FaxName = Name;
            this.dicImage = dicImage;
        }

        private void FaxViewer_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            SetReportForm();
            try
            {
                FaxInfo faxInfo = new FaxInfo();
                FaxService faxService = new FaxService(faxInfo.LinkID, faxInfo.SECRET_KEY);
                faxService.IsTest = ComnEtcFunc.IsFaxTest;
                double dBalance = faxService.GetPartnerBalance(faxInfo.CorpNo);
                string sMSG = string.Empty;
                if (dBalance > 1000)
                {
                    LblBalanceAmt.ForeColor = Color.Blue;
                }
                else
                {
                    sMSG = " ( 포인트 충전이 필요합니다. ) ";
                    LblBalanceAmt.ForeColor = Color.Red;
                }

                LblBalanceAmt.Text = string.Format("{0}P{1}", dBalance, sMSG);
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SetReportForm()
        {
            Cursor = Cursors.WaitCursor;
            switch (FaxName)
            {
                case "RptMesuring":
                    RptMesuring rptMesure = new RptMesuring(dt.Rows[0], dicImage);
                    rptMesure.DataSource = dt;
                    documentViewer1.DocumentSource = rptMesure;
                    rptMesure.CreateDocument();
                    break;
                case "RptMesuring2":
                    RptMesuring2 rptMesure2 = new RptMesuring2(dt.Rows[0], dicImage);
                    rptMesure2.DataSource = dt;
                    documentViewer1.DocumentSource = rptMesure2;
                    rptMesure2.CreateDocument();
                    break;
                case "RptWebFax":
                    RptWebFax rptWebFax = new RptWebFax(dt.Rows[0], dicImage);
                    rptWebFax.DataSource = dt;
                    documentViewer1.DocumentSource = rptWebFax;
                    rptWebFax.CreateDocument();
                    break;
            }
            Cursor = Cursors.Default;
        }

        private void BtnCloseFax_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnSendFax_Click(object sender, EventArgs e)
        {
            string sJunpyoID = dt.Rows[0]["JUNPYOID"]?.ToString();
            string sDealerNM = dt.Rows[0]["DEALER_NM"]?.ToString();
            string sFaxNo = dt.Rows[0]["FAX"]?.ToString();
            //string receiverNum = dt.Rows[0]["FAX"].ToString(); - DB에서 가져온 팩스번호
            if (MessageBox.Show(sDealerNM + "에 계근표를 팩스전송을 진행하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string strFileName = null;
                using (RptWebFax rptMesure = (RptWebFax)documentViewer1.DocumentSource)
                {
                    //rptMesure.DataSource = dt;

                    DirectoryInfo directory = new DirectoryInfo(Application.StartupPath + @"\" + "ticket");
                    if (!directory.Exists)
                        directory.Create();
                    
                    //Guid new_guid = Guid.NewGuid();
                    strFileName = directory + @"\" + "temp.jpg";
                    rptMesure.ExportToImage(strFileName);
                }

                // 발신번호
                string senderNum = "055-345-1295";

                // 수신번호
                string receiverNum = sFaxNo;

                // 수신자명
                string receiverName = sDealerNM;

                //접수번호 직접 작성시 기재
                string requestNum = null;

                DateTime? reserveDT = null;

                string receiptNum = ComnEtcFunc.SendFexPopbil(senderNum, receiverNum, receiverName, strFileName,
                                                                reserveDT, faxTitle, requestNum);
                
                File.Delete(strFileName);

                /*
                 * 2021-01-04
                 * 웹팩스 로그로직 추가
                 */
                try
                {
                    Cursor = Cursors.WaitCursor;

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Add("JUNPYOID", dt.Rows[0]["JUNPYOID"]?.ToString());
                    dicParams.Add("WEB_FAX_NO", receiptNum);
                    dicParams.Add("FAX_SND_NO", sFaxNo);
                    dicParams.Add("PGM_ID", "AccMeasureDev");
                    dicParams.Add("REG_ID", FmMainToolBar2.UserID);
                    dicParams.Add("REG_DT", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO WEB_FAX_LOG ");
                    strSql.AppendLine("           ( JUNPYOID ");
                    strSql.AppendLine("           , WEB_FAX_NO ");
                    strSql.AppendLine("           , FAX_SND_NO ");
                    strSql.AppendLine("           , PGM_ID ");
                    strSql.AppendLine("           , REG_ID ");
                    strSql.AppendLine("           , REG_DT ) ");
                    strSql.AppendLine("     VALUES( @JUNPYOID ");
                    strSql.AppendLine("           , @WEB_FAX_NO ");
                    strSql.AppendLine("           , @FAX_SND_NO ");
                    strSql.AppendLine("           , @PGM_ID ");
                    strSql.AppendLine("           , @REG_ID ");
                    strSql.AppendLine("           , @REG_DT ) ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    foreach(KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                    }
                    cmd.ExecuteNonQuery();

                    Cursor = Cursors.Default;

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;

                    if (!string.IsNullOrEmpty(receiptNum))
                    {
                        string sMSG = string.Format("거래처명 : {0}" +
                            "\r\nFAX 접수번호 : {1}" +
                            "\r\nFAX 전송이 완료되었습니다.", sDealerNM, receiptNum);
                        XtraMessageBox.Show(sMSG);
                        Dispose();
                    }
                }
                catch(Exception ex)
                {
                    Cursor = Cursors.Default;
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    if (!string.IsNullOrEmpty(receiptNum))
                    {
                        XtraMessageBox.Show("팩스는 접수되었으나 LOG는 남기지 못하였습니다." +
                            "\r\n", ex.Message);
                    }
                    else
                    {
                        XtraMessageBox.Show(ex.Message);
                    }
                }
            }
        }
    }
}