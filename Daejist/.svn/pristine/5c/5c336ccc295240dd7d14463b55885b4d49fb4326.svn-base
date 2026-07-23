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
using DevExpress.XtraSplashScreen;
using MySql.Data.MySqlClient;

namespace AccAdm
{
    public partial class SYS010F00 : DevExpress.XtraEditors.XtraForm
    {
        public SYS010F00()
        {
            InitializeComponent();
        }

        private void SYS010F00_Load(object sender, EventArgs e)
        {
            DateFrom.EditValue = DateTime.Today;
            DateTo.EditValue = DateTime.Today;

            setComboTblNm();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void setComboTblNm()
        {
            string tbl = "acbill"
                         + ",acbunf"
                         + ",acbunf2"
                         //+ ",acc_acc_cd"
                         //+ ",acc_acnt_cd"
                         //+ ",acc_acnt_cd_temp"
                         //+ ",acc_auto_slip"
                         //+ ",acc_card_aprv"
                         //+ ",acc_card_cd"
                         //+ ",acc_cash_rcpt"
                         + ",acc_crt_seq"
                         + ",acc_dealer_cd"
                         //+ ",acc_dealer_cd_backup"
                         //+ ",acc_dealer_cd_backup20210120"
                         + ",acc_dealer_history"
                         //+ ",acc_dept_cd"
                         //+ ",acc_incm_cd"
                         //+ ",acc_mgmt_cd"
                         //+ ",acc_over_carried"
                         //+ ",acc_pay_alw_refr"
                         //+ ",acc_pay_back"
                         //+ ",acc_pay_cd"
                         //+ ",acc_prnote_cd"
                         //+ ",acc_prnote_dealt"
                         //+ ",acc_prnote_mgt"
                         //+ ",acc_rmk_cd"
                         //+ ",acc_slip_det"
                         //+ ",acc_slip_evdn"
                         //+ ",acc_slip_mst"
                         //+ ",acc_slip_type"
                         //+ ",acc_tax_dtl"
                         //+ ",acc_tax_mgt"
                         //+ ",acc_withhold_tax"
                         + ",acjanf"
                         //+ ",acjanf_temp"
                         + ",acmstf"
                         //+ ",acmstf_tem"
                         //+ ",actopf"
                         + ",actran"
                         //+ ",actran_backup"
                         //+ ",actran_backup_20200721"
                         //+ ",actran_backup_20200805"
                         //+ ",actran_backup_20210113"
                         //+ ",actran_hist"
                         //+ ",actran_temp"
                         //+ ",chagamreason"
                         + ",com_base_cd"
                         //+ ",companyinfo"
                         //+ ",equip_cd"
                         + ",equip_cd_history"
                         //+ ",equip_consume_history"
                         //+ ",equip_consume_mgt"
                         //+ ",equip_tool_mgt"
                         //+ ",equip_tool_mgt_history"
                         //+ ",gw_bbs_customer"
                         //+ ",gw_bbs_freeboard"
                         //+ ",gw_bbs_manageconference"
                         //+ ",gw_bbs_manageprogress"
                         //+ ",gw_bbs_managepropose"
                         //+ ",gw_bbs_notice"
                         //+ ",gw_bbs_prodconference"
                         //+ ",gw_bbs_prodprogress"
                         //+ ",gw_bbs_prodpropose"
                         //+ ",gw_bbs_salesconference"
                         //+ ",gw_bbs_salesprogress"
                         //+ ",gw_bbs_salespropose"
                         //+ ",gw_board_config"
                         //+ ",gw_board_direct"
                         //+ ",gw_board_file"
                         //+ ",gw_board_media"
                         //+ ",gw_board_notice"
                         //+ ",gw_board_photo"
                         //+ ",gw_board_qna"
                         //+ ",gw_board_recruit"
                         //+ ",gw_comment_direct"
                         //+ ",gw_comment_freeboard"
                         //+ ",gw_comment_media"
                         //+ ",gw_comment_notice"
                         //+ ",gw_comment_photo"
                         //+ ",gw_comment_qna"
                         //+ ",gw_comment_recruit"
                         //+ ",gw_config_bbs"
                         //+ ",gw_dept"
                         //+ ",gw_file_bbs"
                         //+ ",gw_file_popup"
                         //+ ",gw_job"
                         //+ ",gw_member"
                         //+ ",gw_popup"
                         //+ ",gw_scm"
                         //+ ",gw_sequence"
                         //+ ",gw_zip"
                         //+ ",hr_absent"
                         + ",hr_cert_empt"
                         + ",hr_cert_empt_history"
                         + ",hr_emp_academic"
                         + ",hr_emp_basis"
                         + ",hr_emp_career"
                         //+ ",hr_emp_cert"
                         //+ ",hr_emp_family"
                         //+ ",hr_emp_military"
                         + ",hr_emp_personal"
                         + ",hr_emp_salary"
                         + ",income_stl_dtl"
                         + ",income_stl_mgt"
                         + ",inlist"
                         //+ ",inlist_backup_20210113"
                         //+ ",inlist_log"
                         //+ ",inlist_temp"
                         + ",ipchulgo"
                         //+ ",ipchulgo_backup_20210113"
                         //+ ",ipchulgo_temp"
                         + ",jajae"
                         + ",jajaehis"
                         //+ ",junpyo"
                         //+ ",junpyo_mig"
                         //+ ",junpyo_mig_step2"
                         //+ ",kyejung"
                         + ",lengths"
                         //+ ",mab_acc_period"
                         //+ ",mab_acc_stnd_rmk"
                         //+ ",mab_account"
                         //+ ",mab_auto_slip_type"
                         //+ ",mab_bankaccno"
                         //+ ",mab_card"
                         //+ ",mab_crt_seq"
                         //+ ",mab_dealer_info"
                         //+ ",mab_gw_html"
                         //+ ",mab_help"
                         //+ ",mab_incm_r"
                         //+ ",mab_mention"
                         //+ ",mab_mnt_sub"
                         //+ ",mab_payable_bill"
                         //+ ",mab_prt_form"
                         //+ ",mad_auto_slip"
                         //+ ",mad_dpwd_mgt"
                         //+ ",mad_exps_rsltn_det"
                         //+ ",mad_exps_rsltn_mst"
                         //+ ",mad_pay_back"
                         //+ ",mad_slip_clo"
                         //+ ",mad_slip_det"
                         //+ ",mad_slip_evd"
                         //+ ",mad_slip_mst"
                         //+ ",mad_slip_tot"
                         + ",make_1"
                         + ",make_2"
                         + ",make_3"
                         + ",make_4"
                         + ",make_5"
                         //+ ",make_6"
                         + ",make_7"
                         + ",make_expense"
                         + ",make_m"
                         + ",make_s"
                         //+ ",mam_withtax_clo"
                         //+ ",mat_cash_rcpt"
                         //+ ",mat_nts_tax"
                         //+ ",mat_tax"
                         //+ ",mat_tax_dtl"
                         //+ ",mat_with_tax"
                         //+ ",may_carr_for"
                         //+ ",mbf_bugt_assn_det"
                         //+ ",mbf_bugt_assn_req"
                         //+ ",mbf_bugt_assum_mgt"
                         //+ ",mbf_bugt_cd"
                         //+ ",mbf_bugt_clo"
                         //+ ",mbf_bugt_dept_aprv"
                         //+ ",mbf_bugt_dept_mkup"
                         //+ ",mbf_bugt_dept_mkup_dtl"
                         //+ ",mbf_bugt_exec_mkup"
                         //+ ",meaipsiljuk"
                         + ",mesure_ispt_info"
                         + ",mesure_opn_history"
                         + ",mesuring"
                         //+ ",mesuring_backup_20210113"
                         + ",mesuring_seq"
                         //+ ",mesuring_temp"
                         //+ ",plan_1"
                         //+ ",qim_claim"
                         //+ ",rep_stock_manage"
                         + ",salemaechul"
                         + ",salemaeip"
                         + ",sugmf"
                         + ",taxf"
                         //+ ",temp11"
                         //+ ",temp22"
                         //+ ",temp33"
                         + ",tenter"
                         + ",web_fax_log"
                         //+ ",zpgmaut"
                         //+ ",zpgmlst"
                         + ",zsys_log"
                         //+ ",zsys_version"
                         + ",zusrlst";
                         //+ ",zusrskin";

            string[] arr = tbl.Split(',');

            for(int i = 0; i < arr.Length; i++)
            {
                CboTbNm.Properties.Items.Add(arr[i]);
            }
        }

        private void BtnMig_Click(object sender, EventArgs e)
        {
            string sTbl = CboTbNm.EditValue?.ToString();

            if (string.IsNullOrEmpty(sTbl))
                return;

            string strConn = "Server=192.168.0.202;Database=daejierp;Uid=daeji;Pwd=dj2019;Charset=UTF8;allow user variables=true;Allow Zero Datetime=True;";//실서버 admin 122
            string json = string.Empty;

            DataSet ds = new DataSet();
            DataTable dtMysql = new DataTable();
            StringBuilder strSql = new StringBuilder();
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            Cursor = Cursors.WaitCursor;

            try
            {
                

                MySqlConnection dbCon = new MySqlConnection(strConn);

                dbCon.Open();

                strSql.Clear();
                strSql.AppendLine(" SELECT * ");
                strSql.AppendLine("   FROM "+ sTbl);

                MySqlDataAdapter adpt = new MySqlDataAdapter(strSql.ToString(), dbCon);

                MySqlTransaction dbTran = adpt.SelectCommand.Transaction;

                adpt.Fill(ds);

                dtMysql = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }

            if(dtMysql != null)
            {
                //mysql row 개수
                int dtMyCnt = dtMysql.Rows.Count;
                //한번에 전달할 row 개수
                int transCnt = 1000;
                
                if(dtMyCnt < transCnt)
                {
                    json = ComnEtcFunc.DataTableToJsonObj(dtMysql);

                    if (!string.IsNullOrEmpty(json))
                    {
                        dicParams.Clear();
                        dicParams.Add("CMD", sTbl);
                        dicParams.Add("JSON", json);

                        DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_MIG", dicParams);

                        if (dt != null)
                        {
                            XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());
                        }
                    }
                }
                else
                {
                    //시작행
                    int sRowCnt = 0;
                    while (true)
                    {
                        DataTable dtTemp = dtMysql.Clone();

                        if (dtMyCnt - transCnt > 0)
                        {
                            for (int i= sRowCnt; i< sRowCnt+transCnt; i++)
                            {
                                dtTemp.ImportRow(dtMysql.Rows[i]);
                            }

                            json = ComnEtcFunc.DataTableToJsonObj(dtTemp);

                            if (!string.IsNullOrEmpty(json))
                            {
                                dicParams.Clear();
                                dicParams.Add("CMD", sTbl);
                                dicParams.Add("JSON", json);

                                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_MIG", dicParams);

                                if (dt != null && dt.Rows.Count > 0 && dt.Rows[0]["RESULT"].Equals("0"))
                                {
                                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());
                                    break;
                                }
                            }

                            //남은행 빼주기
                            dtMyCnt -= transCnt;
                            //시작행 더해주기
                            sRowCnt += transCnt;
                        }
                        else
                        {
                            for (int i = sRowCnt; i < sRowCnt + dtMyCnt; i++)
                            {
                                dtTemp.ImportRow(dtMysql.Rows[i]);
                            }

                            json = ComnEtcFunc.DataTableToJsonObj(dtTemp);

                            if (!string.IsNullOrEmpty(json))
                            {
                                dicParams.Clear();
                                dicParams.Add("CMD", sTbl);
                                dicParams.Add("JSON", json);

                                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_MIG", dicParams);

                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());
                                }
                            }

                            break;
                        }
                    }
                }

            }

            Cursor = Cursors.Default;
        }
    }
}