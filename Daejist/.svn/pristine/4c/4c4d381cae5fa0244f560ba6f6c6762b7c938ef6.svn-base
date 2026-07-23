using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccAdm
{

    public partial class ReportViewer : Form
    {

        private static DataTable dt;
        private static DataTable dt2;
        private string ReportName;
        private ArrayList ARRAY_LST;
        private List<string> _LST;
        DataRow _Row;
        Dictionary<string, Image> _dicParam;

        public ReportViewer()
        {
            InitializeComponent();
        }

        public ReportViewer(DataTable dtM, DataTable dtD, string Name)
        {
            InitializeComponent();
            dt = dtM;
            dt2 = dtD;
            ReportName = Name;
        }
        
        public ReportViewer(DataTable dt1, string Name)
        {
            InitializeComponent();
            dt = dt1;
            ReportName = Name;
        }
        public ReportViewer(DataTable dtShredder, string Name, string date)
        {
            
        }

        public ReportViewer(DataTable dt1, ArrayList arr, string Name)
        {
            InitializeComponent();
            dt = dt1;
            ReportName = Name;
            ARRAY_LST = arr;
        }

        public ReportViewer(DataTable dt1, List<string> arr, string Name)
        {
            InitializeComponent();
            dt = dt1;
            ReportName = Name;
            _LST = arr;
        }

        public ReportViewer(DataTable dt1, string sReportName, Dictionary<string, Image> dicParam)
        {
            InitializeComponent();
            dt = dt1;
            ReportName = sReportName;
            _dicParam = dicParam;
        }

        private void AccRptDealer_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            Cursor = Cursors.WaitCursor;
            SetReportForm();
            //switch (ReportName)
            //{
            //    case "DealerCdReport":
            //        RptDealerCd report = new RptDealerCd();
            //        report.DataSource = dt;
            //        documentViewer1.DocumentSource = report;
            //        report.CreateDocument();
            //        break;

            //    case "DailyMonthlyForm":
            //        RptDailyMonthly reportDailyMonth = new RptDailyMonthly();
            //        reportDailyMonth.DataSource = dt;
            //        documentViewer1.DocumentSource = reportDailyMonth;
            //        reportDailyMonth.CreateDocument();
            //        break;
            //    case "RptCertOfEmp":
            //        RptCertOfEmp rptCert = new RptCertOfEmp();
            //        rptCert.DataSource = dt;
            //        documentViewer1.DocumentSource = rptCert;
            //        rptCert.CreateDocument();
            //        break;
            //}
            Cursor = Cursors.Default;
        }

        private void SetReportForm()
        {
            Cursor = Cursors.WaitCursor;
            switch (ReportName)
            {
                case "DealerCdReport":
                    RptDealerCd report = new RptDealerCd();
                    report.DataSource = dt;
                    documentViewer1.DocumentSource = report;
                    report.CreateDocument();
                    break;

                case "DailyMonthlyForm":
                    RptDailyMonthly reportDailyMonth = new RptDailyMonthly();
                    reportDailyMonth.DataSource = dt;
                    documentViewer1.DocumentSource = reportDailyMonth;
                    reportDailyMonth.CreateDocument();
                    break;
                case "RptCertOfEmp":
                    RptCertOfEmp rptCert = new RptCertOfEmp();
                    rptCert.DataSource = dt;
                    documentViewer1.DocumentSource = rptCert;
                    rptCert.CreateDocument();
                    break;
                case "RptMesuring":
                    RptMesuring rptMesure = new RptMesuring();
                    rptMesure.DataSource = dt;
                    documentViewer1.DocumentSource = rptMesure;
                    rptMesure.CreateDocument();
                    break;
                case "RptVoucher":
                    RptVoucher rptVoucher = new RptVoucher(dt, ARRAY_LST);
                    rptVoucher.DataSource = dt;
                    documentViewer1.DocumentSource = rptVoucher;
                    rptVoucher.CreateDocument();
                    break;
                case "RptWebFax":
                    RptWebFax rptWebFax = new RptWebFax(dt.Rows[0], _dicParam);
                    //rptWebFax.DataSource = dt;
                    documentViewer1.DocumentSource = rptWebFax;
                    rptWebFax.CreateDocument();
                    break;
                case "RptTaxBill":
                    RptTaxBill rptTax = new RptTaxBill(dt);
                    RptTaxBill2 rptTax2 = new RptTaxBill2(dt);
                    //rptTax.DataSource = dt;
                    rptTax.CreateDocument();
                    rptTax2.CreateDocument();

                    foreach (DevExpress.XtraPrinting.Page page in rptTax2.Pages)
                    {
                        rptTax.Pages.Add(page);
                    }
                    documentViewer1.DocumentSource = rptTax;
                    break;
                case "RptAcc":
                    RptAcc RE02S = new RptAcc(dt, dt2);
                    RE02S.DataSource = dt2;
                    documentViewer1.DocumentSource = RE02S;
                    RE02S.CreateDocument();
                    break;
                case "RptMesuring2":
                    RptMesuring2 rptMesure2 = new RptMesuring2(dt.Rows[0], _dicParam);
                    rptMesure2.DataSource = dt;
                    documentViewer1.DocumentSource = rptMesure2;
                    rptMesure2.CreateDocument();
                    break;
                case "RptInOutSlip":
                    RptInOutSlip rptInOutSlip = new RptInOutSlip(_LST, dt);
                    rptInOutSlip.DataSource = dt;
                    documentViewer1.DocumentSource = rptInOutSlip;
                    rptInOutSlip.CreateDocument();
                    break;
                case "RptDealerImage":
                    RptDealerImage rptDealerImage = new RptDealerImage(dt);
                    rptDealerImage.DataSource = dt;
                    documentViewer1.DocumentSource = rptDealerImage;
                    rptDealerImage.CreateDocument();
                    break;
            }
            Cursor = Cursors.Default;
        }

        private void bbiPrintDirect_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }
    }
}
