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
using MySql.Data.MySqlClient;

namespace AccAdm
{
    public partial class PopUpEvdn : DevExpress.XtraEditors.XtraForm
    {
        public DataTable dtEvdn { get; set; }
        public string sDealerCd { get; set; }
        public string sIncmrCd { get; set; }
        public DataTable dtEvdnPop { get; set; }

        public PopUpEvdn()
        {
            InitializeComponent();
        }

        private void PopUpEvdn_Load(object sender, EventArgs e)
        {
            DataTable dtEvdnGb = GetLookUpData("1", "N", "");

            ComLib.ComGrid.SetLookUpEdit(LkupEvdnGb, dtEvdnGb, "CD", "NM", "Y");

            RepositoryItemGridLookUpEdit EvdnLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(EvdnLkup, dtEvdnGb, GridEvdnInfo, GridColEvdnGb, "CD", "NM", "");

            DateEditEvdnYmd.EditValueChanged -= DateEditEvdnYmd_EditValueChanged;
            DateEditEvdnYmd.EditValue = System.DateTime.Now;
            DateEditEvdnYmd.EditValueChanged += DateEditEvdnYmd_EditValueChanged;

            //GridViewEvenInfo.FocusedRowChanged -= GridViewEvenInfo_FocusedRowChanged;
            dtEvdnPop = dtEvdn.Copy();
            GridEvdnInfo.DataSource = dtEvdnPop;
            //GridViewEvenInfo.FocusedRowChanged += GridViewEvenInfo_FocusedRowChanged;
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '****' AS CD ");
                strSql.AppendLine("      , '전체' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine("  WHERE CD_GB = 'EVDN_GB' ");
                strSql.AppendLine("    AND USE_YN = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.ACC_CD AS CD ");
                strSql.AppendLine("      , A.ACC_NM AS NM  ");
                strSql.AppendLine("   FROM ACC_ACC_CD A ");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT IDT_NO AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A ");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine(" , COM_NM AS NM ");
                strSql.AppendLine(" FROM COM_BASE_CD ");
                strSql.AppendLine(" WHERE CD_GB = 'EVDN_GB' ");
            }
            else if (sGb.Equals("5"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine(" , COM_NM AS NM ");
                strSql.AppendLine(" FROM COM_BASE_CD ");
                strSql.AppendLine(" WHERE CD_GB = 'MGMT_GB' ");
            }
            else if (sGb.Equals("8"))
            {
                strSql.AppendLine(" SELECT CD ");
                strSql.AppendLine("      , NM ");
                strSql.AppendLine("   FROM V_MGMT_CD ");
                strSql.AppendLine(" WHERE MGMT_GB = '" + sParam + "' ");
            }
            else if (sGb.Equals("9"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine(" WHERE CD_GB = 'BANK_CD' ");
            }
            else if (sGb.Equals("10"))
            {
                strSql.AppendLine(" SELECT COM_CD AS CD ");
                strSql.AppendLine("      , COM_NM AS NM ");
                strSql.AppendLine("   FROM COM_BASE_CD ");
                strSql.AppendLine(" WHERE CD_GB = 'EVDN_GB' ");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY CD");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            GridViewEvenInfo.FocusedRowChanged -= GridViewEvenInfo_FocusedRowChanged;
            GridViewEvenInfo.AddNewRow();
            string sGb = LkupEvdnGb.EditValue.ToString();
            SetGridValue(sGb);
            GridViewEvenInfo.FocusedRowChanged += GridViewEvenInfo_FocusedRowChanged;
        }
        
        private void InitCtrlByEvdnGb(string sGb)
        {
            TxtEvdnAmt.Enabled = false;
            TxtSuplAmt.Enabled = false;
            TxtVatAmt.Enabled = false;
            BtneIdtNo.Enabled = false;
            TxtDealerNm.Enabled = false;
            BtneIncmrIdtNo.Enabled = false;
            TxtIncmrNm.Enabled = false;
            BtnEvdn.Visible = true;
            BtnAdd.Enabled = false;

            if (sGb.Equals("01") || sGb.Equals("02"))
            {
                SetCtrlClear();

                BtneIdtNo.Enabled = true;
                TxtDealerNm.Enabled = true;
                TxtDealerNm.Focus();

                BtnAdd.Enabled = true;
                // 세금계산서 내역 팝업
            }
            else if (sGb.Equals("05"))
            {
                SetCtrlClear();
                BtneIncmrIdtNo.Enabled = true;
                TxtIncmrNm.Enabled = true;
                TxtIncmrNm.Focus();
                // 원천영수증 내역 팝업
            }
            else if (sGb.Equals("03"))
            {
                SetCtrlClear();
            }
            else if (sGb.Equals("04"))
            {
                SetCtrlClear();
            }

            if (sGb.Equals("06") || sGb.Equals("07"))
            {
                BtnEvdn.Visible = false;
                BtnAdd.Enabled = true;
                TxtEvdnAmt.Enabled = true;

                

                SetCtrlClear();
                SetGridValue(sGb);

                TxtEvdnAmt.Focus();
            }
            else
            {
                BtnEvdn.Text = LkupEvdnGb.Text;
            }
        }

        private void SetCtrlClear()
        {
            GridViewEvenInfo.FocusedRowChanged -= GridViewEvenInfo_FocusedRowChanged;
            GridViewEvenInfo.AddNewRow();
            GridViewEvenInfo.FocusedRowChanged += GridViewEvenInfo_FocusedRowChanged;

            BtneIdtNo.EditValue = "";
            TxtDealerNm.EditValue = "";
            BtneIncmrIdtNo.EditValue = "";
            TxtIncmrNm.EditValue = "";
            TxtEvdnAmt.EditValue = 0;
            TxtSuplAmt.EditValue = 0;
            TxtVatAmt.EditValue = 0;
            TxtNote.EditValue = "";
            BtneEmpId.EditValue = "";
            TxtEmpNm.EditValue = "";
        }

        private void SetGridValue(string sGb)
        {
            GridViewEvenInfo.SetFocusedRowCellValue("EVDN_SEQ", GridViewEvenInfo.RowCount);
            GridViewEvenInfo.SetFocusedRowCellValue("EVDN_GB", sGb);
            GridViewEvenInfo.SetFocusedRowCellValue("EVDN_YMD", Convert.ToString(DateEditEvdnYmd.EditValue));
            GridViewEvenInfo.SetFocusedRowCellValue("EVDN_AMT", TxtEvdnAmt.EditValue);
            GridViewEvenInfo.SetFocusedRowCellValue("SUPL_AMT", TxtSuplAmt.EditValue);
            GridViewEvenInfo.SetFocusedRowCellValue("VAT_AMT", TxtVatAmt.EditValue);
            GridViewEvenInfo.SetFocusedRowCellValue("EMP_ID", Convert.ToString(BtneEmpId.EditValue));
            GridViewEvenInfo.SetFocusedRowCellValue("EMP_NM", Convert.ToString(TxtEmpNm.EditValue));
            GridViewEvenInfo.SetFocusedRowCellValue("NOTE", Convert.ToString(TxtNote.EditValue));
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            GridViewEvenInfo.DeleteRow(GridViewEvenInfo.FocusedRowHandle);
        }

        private void BtnComfirm_Click(object sender, EventArgs e)
        {
            dtEvdnPop = (DataTable)GridEvdnInfo.DataSource;

            this.DialogResult = DialogResult.OK;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }


        private void LkupEvdnGb_EditValueChanged(object sender, EventArgs e)
        {
            InitCtrlByEvdnGb(LkupEvdnGb.EditValue.ToString());
        }

        private void BtnEvdn_Click(object sender, EventArgs e)
        {
            GridViewEvenInfo.FocusedRowChanged -= GridViewEvenInfo_FocusedRowChanged;
            if(Convert.ToString(GridViewEvenInfo.GetFocusedRowCellValue("EVDN_AMT")).Replace(",", "").Equals("0"))
            GridViewEvenInfo.DeleteRow(GridViewEvenInfo.FocusedRowHandle);
            GridViewEvenInfo.FocusedRowChanged += GridViewEvenInfo_FocusedRowChanged;

            string sGb = LkupEvdnGb.EditValue.ToString();

            if (sGb.Equals("01") || sGb.Equals("02"))
            {
                PopUpVatInfo frm = new PopUpVatInfo();

                frm.iTaxGb = sGb.Equals("01") ? 0 : 1;
                frm.sDealerNm = TxtDealerNm.Text;
                frm.sIdtNo = Convert.ToString(BtneIdtNo.EditValue);

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    if(frm.dtResult.Rows.Count > 0)
                    {
                        SetVatInfo(sGb, frm.dtResult);
                    }
                }
                else
                {
                    return;
                }
            }
            else if (sGb.Equals("03"))
            {
                
            }
            else if (sGb.Equals("04"))
            {

            }
            else if (sGb.Equals("05"))
            {
                PopUpWithHoldInfo frm = new PopUpWithHoldInfo();

                frm.sIncmrIdtNo = Convert.ToString(BtneIncmrIdtNo.EditValue);
                frm.sIncmrNm = TxtIncmrNm.Text.Trim();
                frm.sIncmrCd = sIncmrCd;

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    if(frm.dtResult.Rows.Count > 0)
                    {
                        SetWithHoldingInfo(sGb, frm.dtResult);
                    }
                }
                else
                {
                    return;
                }
            }
        }

        private void SetVatInfo(string sGb, DataTable dt)
        {
            DataTable dtOld = new DataTable();
            dtOld = (DataTable)GridEvdnInfo.DataSource;
            string sBillKey = string.Empty;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sBillKey = Convert.ToString(dt.Rows[i]["BILL_KEY"]);

                DataRow[] dr = dtOld.Select("EVDN_KEY = '" + sBillKey + "'");

                if (dr.Length == 0)
                {
                    GridViewEvenInfo.AddNewRow();
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_SEQ", GridViewEvenInfo.RowCount);
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_GB", sGb);
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_YMD", Convert.ToString(dt.Rows[i]["BILL_ISSUE_YMD"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_AMT", Convert.ToDouble(dt.Rows[i]["TOT_AMT"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("SUPL_AMT", Convert.ToDouble(dt.Rows[i]["SUPL_AMT"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("VAT_AMT", Convert.ToDouble(dt.Rows[i]["VAT_AMT"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("DEALER_CD", Convert.ToString(dt.Rows[i]["DEALER_CD"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("IDT_NO", Convert.ToString(dt.Rows[i]["IDT_NO"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("DEALER_NM", Convert.ToString(dt.Rows[i]["DEALER_NM"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_KEY", Convert.ToString(dt.Rows[i]["BILL_KEY"]));
                }
            }

            GridViewEvenInfo_FocusedRowChanged(null, null);
        }

        private void SetWithHoldingInfo(string sGb, DataTable dt)
        {
            DataTable dtOld = new DataTable();
            dtOld = GridViewEvenInfo.DataSource as DataTable;
            string sEvenKey = string.Empty;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sEvenKey = Convert.ToString(dt.Rows[i]["EVDN_KEY"]);

                DataRow[] dr = dtOld.Select("EVDN_KEY = sEvenKey");

                if (dr.Length == 0)
                {
                    GridViewEvenInfo.AddNewRow();
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_SEQ", GridViewEvenInfo.RowCount);
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_GB", sGb);
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_YMD", Convert.ToString(dt.Rows[i]["PMNT_YMD"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_AMT", Convert.ToDouble(dt.Rows[i]["WITHTAX_TOT"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("SUPL_AMT", 0);
                    GridViewEvenInfo.SetFocusedRowCellValue("VAT_AMT", 0);
                    GridViewEvenInfo.SetFocusedRowCellValue("INCMR_IDT_NO", Convert.ToString(dt.Rows[i]["INCMR_IDT_NO"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("INCMR_NM", Convert.ToString(dt.Rows[i]["INCMR_NM"]));
                    GridViewEvenInfo.SetFocusedRowCellValue("EVDN_KEY", Convert.ToString(dt.Rows[i]["EVDN_KEY"]));
                }
            }
        }

        public DataRow drDealerInfo;

        private void BtneIdtNo_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataTable dtDealerInfo = ComLib.ClsFunc.GetAccDefaultInfo(Convert.ToString(BtneIdtNo.EditValue).Replace("_", ""));

            if (dtDealerInfo.Rows.Count == 1)
            {
                drDealerInfo = dtDealerInfo.Rows[0];
                BtneIdtNo.EditValue = Convert.ToString(drDealerInfo["IDT_NO"]);
                TxtDealerNm.EditValue = Convert.ToString(drDealerInfo["DEALER_NM"]);
                sDealerCd = Convert.ToString(drDealerInfo["DEALER_CD"]);

                BtnEvdn.PerformClick();
            }
            else
            {
                PopUpDealerCd frm = new PopUpDealerCd();
                frm.sIdtNoByVal = Convert.ToString(BtneIdtNo.EditValue).Replace("_", "");
                frm.sDealerNmByVal = TxtDealerNm.Text.Trim();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    drDealerInfo = frm.drResult;
                    BtneIdtNo.EditValue = Convert.ToString(drDealerInfo["IDT_NO"]);
                    TxtDealerNm.EditValue = Convert.ToString(drDealerInfo["DEALER_NM"]);
                    sDealerCd = Convert.ToString(drDealerInfo["DEALER_CD"]);

                    BtnEvdn.PerformClick();
                }
                else
                {
                    TxtDealerNm.Focus();
                }
            }
        }

        public DataRow drIncmRInfo;

        private void BtneIncmrIdtNo_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataTable dtIncmrInfo = ComLib.ClsFunc.GetIncmRByIdtNo(Convert.ToString(TxtIncmrNm.EditValue).Trim());

            if (dtIncmrInfo.Rows.Count == 1)
            {
                drDealerInfo = dtIncmrInfo.Rows[0];
                BtneIncmrIdtNo.EditValue = Convert.ToString(drIncmRInfo["INCMR_IDT_NO"]);
                TxtIncmrNm.EditValue = Convert.ToString(drIncmRInfo["INCMR_NM"]);
                sIncmrCd = Convert.ToString(drIncmRInfo["INCMR_CD"]);
            }
            else
            {
                PopUpIncmrCd frm = new PopUpIncmrCd();
                frm.sIncmrNmByVal = TxtIncmrNm.Text.Trim();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    drDealerInfo = frm.drResult;
                    BtneIncmrIdtNo.EditValue = Convert.ToString(drIncmRInfo["INCMR_IDT_NO"]);
                    TxtIncmrNm.EditValue = Convert.ToString(drIncmRInfo["INCMR_NM"]);
                    sIncmrCd = Convert.ToString(drIncmRInfo["INCMR_CD"]);
                }
                else
                {
                    TxtIncmrNm.Focus();
                }
            }
        }

        private void TxtIncmrNm_Leave(object sender, EventArgs e)
        {
            BtneIncmrIdtNo_ButtonClick(null, null);
        }

        public DataRow drEmpId;

        private void BtneEmpId_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataTable dtEmpInfo = ComLib.ClsFunc.GetEmpInfoByEmpId(Convert.ToString(BtneEmpId.EditValue).Trim());

            if (dtEmpInfo.Rows.Count == 1)
            {
                drEmpId = dtEmpInfo.Rows[0];
                BtneEmpId.EditValue = Convert.ToString(drEmpId["EMP_ID"]);
                TxtEmpNm.EditValue = Convert.ToString(drEmpId["EMP_NM"]);
            }
            else
            {
                PopUpEmployeeCd frm = new PopUpEmployeeCd();
                frm.sEmpId = Convert.ToString(BtneEmpId.EditValue).Trim();
                frm.sEmpNm = Convert.ToString(TxtEmpNm.EditValue).Trim();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    drEmpId = frm.drResult;
                    BtneEmpId.EditValue = Convert.ToString(drEmpId["EMP_ID"]);
                    TxtEmpNm.EditValue = Convert.ToString(drEmpId["EMP_NM"]);
                }
                else
                {
                    TxtIncmrNm.Focus();
                }
            }
        }
        
        private void DateEditEvdnYmd_EditValueChanged(object sender, EventArgs e)
        {
            SetGridValue(LkupEvdnGb.EditValue.ToString());
        }

        private void TxtNote_EditValueChanged(object sender, EventArgs e)
        {
            SetGridValue(LkupEvdnGb.EditValue.ToString());
        }

        private void TxtEvdnAmt_EditValueChanged(object sender, EventArgs e)
        {
            SetGridValue(LkupEvdnGb.EditValue.ToString());
        }

        private void repositoryItemDateEdit2_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            string sYmd = e.Value.ToString();
            sYmd = sYmd.Replace("-", "").Substring(0, 8);
            e.DisplayText = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
        }

        private void TxtEvdnAmt_Leave(object sender, EventArgs e)
        {
            if(LkupEvdnGb.EditValue.ToString().Equals("06") || LkupEvdnGb.EditValue.ToString().Equals("07"))
            {
                TxtNote.Focus();
            }
        }

        private void TxtEmpNm_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.KeyValue == 13)
            {
                BtneEmpId_ButtonClick(null, null);
                SetGridValue(LkupEvdnGb.EditValue.ToString());
            }
        }

        private void GridViewEvenInfo_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            LkupEvdnGb.EditValueChanged -= LkupEvdnGb_EditValueChanged;
            DateEditEvdnYmd.EditValueChanged -= DateEditEvdnYmd_EditValueChanged;
            TxtEvdnAmt.EditValueChanged -= TxtEvdnAmt_EditValueChanged;
            TxtNote.EditValueChanged -= TxtNote_EditValueChanged;

            LkupEvdnGb.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("EVDN_GB");
            DateEditEvdnYmd.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("EVDN_YMD");
            TxtEvdnAmt.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("EVDN_AMT");
            TxtSuplAmt.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("SUPL_AMT");
            TxtVatAmt.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("VAT_AMT");
            BtneIdtNo.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("IDT_NO");
            TxtDealerNm.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("DEALER_NM");
            BtneIncmrIdtNo.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("INCMR_IDT_NO");
            TxtIncmrNm.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("INCMR_NM");
            BtneEmpId.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("EMP_ID");
            TxtEmpNm.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("EMP_NM");
            TxtNote.EditValue = GridViewEvenInfo.GetFocusedRowCellValue("NOTE");

            LkupEvdnGb.EditValueChanged += LkupEvdnGb_EditValueChanged;
            DateEditEvdnYmd.EditValueChanged += DateEditEvdnYmd_EditValueChanged;
            TxtEvdnAmt.EditValueChanged += TxtEvdnAmt_EditValueChanged;
            TxtNote.EditValueChanged += TxtNote_EditValueChanged;
        }

        private void TxtDealerNm_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                BtneIdtNo_ButtonClick(null, null);
            }
        }

        private void PopUpEvdn_KeyDown(object sender, KeyEventArgs e)
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
                BtnComfirm_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnEvdn_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
               
            }
        }

        private void GridEvdnInfo_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewEvenInfo_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}