using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Xml;

namespace AccAdm
{
    public partial class PP008F00 : DevExpress.XtraEditors.XtraForm
    {
        string currentPage = "1";  //현재 페이지
        string countPerPage = "1000"; //1페이지당 출력 갯수
        string confmKey = "U01TX0FVVEgyMDIxMDYzMDE2NTEwODExMTM0MzE="; //승인 Key
        string keyword = string.Empty;
        string apiurl = string.Empty;
        public delegate void SendDataHandler(DataRow row);
        public event SendDataHandler DataRowSendEvent;
        public PP008F00()
        {
            InitializeComponent();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                keyword = TxtWord.Text.Trim();
                apiurl = "http://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage=" + currentPage + "&countPerPage=" + countPerPage + "&keyword=" + keyword + "&confmKey=" + confmKey;

                WebClient wc = new WebClient();

                XmlReader read = new XmlTextReader(wc.OpenRead(apiurl));

                DataSet ds = new DataSet();
                ds.ReadXml(read);

                DataRow[] rows = ds.Tables[0].Select();

                if (rows[0]["totalCount"].ToString() != "0")
                {
                    GridRetr.DataSource = ds.Tables[1];
                }

            }
            catch (Exception ex)
            {
            }

            GridViewRetr.Focus();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            DataRowSendEvent?.Invoke(GridViewRetr.GetFocusedDataRow());
            Dispose();
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
                BtnSave.PerformClick();
        }

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSave.PerformClick();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void TxtWord_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
                BtnRetr.Focus();
        }
    }
}