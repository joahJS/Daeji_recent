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
using DevExpress.XtraGrid.Columns;
using ComLib;

namespace AccAdm
{
    public partial class SC010F01 : DevExpress.XtraEditors.XtraForm
    {
        private string PROCEDURE_ID = "DP_SC010F00";

        public delegate void SendDataHandler(string sVal);
        public event SendDataHandler DataRowSendEvent;

        public SC010F01()
        {
            InitializeComponent();
        }

        private void SC010F01_Load(object sender, EventArgs e)
        {
            KeyPreview = true;
            ComnGridFunc.GridStyleBasicSetting(GridViewRetr);
            KPIRetr();
        }

        private void Bt_Save_Click(object sender, EventArgs e)
        {
            KPIUpdate();
        }

        private void Bt_Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        // KPI 성과지표 조회 메서드
        private void KPIRetr()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "KPI_Retr");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr.DataSource = dt;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        // KPI 목표 업데이트 메서드
        private void KPIUpdate()
        {
            try
            {
                DataTable dtKPI = (DataTable)GridRetr.DataSource;

                DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dtKPI);
                DataTable dtMerge = dsSave.Tables[0];

                if (dtMerge == null || dtMerge.Rows.Count == 0)
                {
                    XtraMessageBox.Show("변경된 데이터가 없습니다.");
                    return;
                }

                GridColumn[] Gc = new GridColumn[] { GridColPRATE };
                string errMsg = ComnGridFunc.GridViewEmptyValueCheck(GridViewRetr, Gc);

                if (!string.IsNullOrEmpty(errMsg))
                {
                    XtraMessageBox.Show(errMsg, "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridRetr.Focus();
                    return;
                }

                string JSON1 = ComnEtcFunc.DataTableToJsonObj(dtMerge);

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "KPI_Update");
                dicParams.Add("JSON1", JSON1);
                dicParams.Add("USRCD", FmMainToolBar2.UserID);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        XtraMessageBox.Show(dt.Rows[0]["MSG"].ToString(), "저장");
                        if (dt.Rows[0]["RESULT"].ToString().Equals("1"))
                        {
                            DataRowSendEvent("Refresh");
                            Close();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        #region [ 키 이벤트 ]
        private void SC010F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3) { Bt_Save.PerformClick(); }
            else if (e.KeyCode == Keys.Escape) { }
        } 
        #endregion
    }
}