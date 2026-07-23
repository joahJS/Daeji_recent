using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using ComLib;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraEditors;
using System.Diagnostics;
using System.Threading;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.IO;

namespace AccAdm
{
    public partial class FmMainToolBar2 : DevExpress.XtraEditors.XtraForm
    {
        public FmMainToolBar2()
        {
            InitializeComponent();
        }

        public static DataRow drUser { get; set; }
        public static DataTable dtUserAutInfo { get; set; }
        public static string UserID { get; set; }
        public static string EmpID { get; set; }
        public static string _VERSION_ID;
        public static string SAVE_LAYOUT_LOADING_NAME = "_Layout 저장중...";
        public static PrivateFontSetting _FontSetting;

        private string LoginTime;

        Thread _T_CUR_TIME;
        Thread _T_ALARM;
        Thread _T_ALARMIMG;
        Thread _T_F_ALARM;

        private void FmMainToolBar2_Load(object sender, EventArgs e)
        {
            //
            DBConn.dbCon = DBConn.DbConn();
            SetUserInfo();
            GetCurTime();


            //알람
            LoginTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SetFAlarm();

            openAlarmDelegate = new OpenAlarmDelegate(openAlarmForm);
            GetAlarm();

            changeAlarmImgDelegate = new ChangeAlarmImgDelegate(ChangeAlarmImg);
            SetAlarmImg();

            ComnEtcFunc.LogAPI("접속", ComnEtcFunc.GetLocalIP());
        }

        private void FmMainToolBar2_Shown(object sender, EventArgs e)
        {
            string sUserCd = drUser["USRCD"]?.ToString();

            if (string.IsNullOrEmpty(sUserCd))
            {
                XtraMessageBox.Show("로그인 정보가 없습니다. 다시 로그인 해주세요.");
                return;
            }
            //UserID = ComLib.ClsFunc.GetUserIdOfLoginInfo(sUserCd); //LoginID 저장 -> INSERT,UPDATE 시의 MFY, ENT 저장 위함
            UserID = sUserCd;
            EmpID = drUser["INSANO"]?.ToString();
            /*
             * 2021-01-19
             * 스킨적용로직 추가
             */
            SkinInfo skin = new SkinInfo(UserID);
            skin.SetProgramSkin();

            _FontSetting = new PrivateFontSetting(UserID);
            
            DataTable dt = ComnEtcFunc.GetAuthInfo(sUserCd);
            
            //ShowMenuByAutority(dt);

            dtUserAutInfo = dt;

            BarSubItem[] ArrBigCate = new BarSubItem[] { ToolBarPageBasic, ToolBarPageProd, ToolBarPageInOut, ToolBarPageDevice, ToolBarPageQC, ToolBarPageHR, ToolBarPageACC, ToolBarPageReport, ToolBarBtnSystem };
            BarButtonItem[] ArrSubItem = new BarButtonItem[] { BarbtnItCompanyInfo, BarbtnItDealerCdDev, BarbtnItAccCdMgtDev, BarbtnItPayAlw , BarbtnItPaymentCd , BarbtnItAccountNumber ,
            BarbtnItCreditCardDev, BarbtnItCommonCodeMgt, BarbtnChagamReason, BarbtnItUnitCostMgt, BarbtnItJaJae, BarbtnItSlipType, BarbtnItEquipCdMgt, BarbtnItUserMgt, BarbtnItConsumeCd, BarbtnItStandardCd, BarbtnItDeptCdDev, BarBtnDoctMgt,
            BarbtnItPlan, BarbtnItProdStatus, BarBtnPD04001F01,//, BarBtnPD02001F01
            BarbtnItMeasure, BarbtnItClose, BarbtnItMeasureDevImg, BarbtnItSiteDev, BarbtnMesMngProgram, BarbtnItInComeDailyReport, BarbtnIN07001F01, BatBtnIN09001F01,BarBtnIN11001F01,BarBtnSC010F00,
            BarbtnItEquipMgt, BarbtnItConsInOut, barButtonItem3, BarBtnEquipHisMgt, BarBtnEQ001F00,
            BarbtnItClaim, BarbtnItQimBadMgt,
            BarbtnItPersonnelRecords, BarBtnItCtfctnEmploee, BarBtnGE001F00, BarBtnGE002F00, BarBtnGE003F00, BarBtnSL002F00,BarBtnIN001F02, BarBtnSL001F00, BarBtnSL001F01,
            BarBtnAC13001F01, BarbtnItSlip, BarbtnItJournal, BarbtnItDailyBalance, BarBtnAC07001F01, BarBtnAC06001F01, BatBtnAC08001F01, BarBtnAC09001F01, BarBtnAC10001F01, BarBtnAC11001F01, BarBtnSYS001F01, BarBtnAC14001F02, BarBtnAC16001F01, BarBtnAC17001F01, BarBtnAC18001F01, BarBtnAC19001F00,BarBtnAC20001F00,
            BarbtnItScrab, BarbtnItLstSalePlan, BarbtnItLstSaleResult, BarbtnItDailyMonthlyRpt, BarbtnItPurcPfrm, BarBtnSYS002F01, BatBtnSYS003F01,BarBtnBusinessReport, BarBtnRptApplSystem,BarBtnAllReport,barBtnAccFieldCustom,
            BarBtnIN10001F01, BarBtnIN12001F01,BarBtnSI001F00
            };
            
            string[] ArrSubItemName = new string[ArrSubItem.Length];
            for(int i = 0; i < ArrSubItem.Length; i++)
            {
                ArrSubItem[i].Visibility = BarItemVisibility.Never;
                ArrSubItemName[i] = ArrSubItem[i].Description;
            }

            ComnEtcFunc.SetMenuListByAuth(dt, ToolBarPageBasic, ToolBarPageProd, ToolBarPageInOut, ToolBarPageDevice, ToolBarPageQC, ToolBarPageHR, ToolBarPageACC,
                ToolBarPageReport, ToolBarBtnSystem, ArrSubItem, ArrSubItemName);
        }

        private void SetUserInfo()
        {
            BarStaticExsId.Caption = string.Format("사용자 : {0}", drUser["USRNM"]?.ToString());
            BarStaticVersion.Caption = string.Format("Copyright DaejiSt (Ver {0})", _VERSION_ID);
        }

        private void GetCurTime()
        {
            if (_T_CUR_TIME == null)
            {
                _T_CUR_TIME = new Thread(new ThreadStart(this.GetTime));
            }

            if (_T_CUR_TIME.IsAlive)
            {
                _T_CUR_TIME.Abort();
            }

            _T_CUR_TIME.IsBackground = true;
            _T_CUR_TIME.Start();
        }

        private void GetTime()
        {
            while (true)
            {
                try
                {
                    BatStaticCurTime.Caption = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {

                }
            }
        }

        //private void ShowMenuByAutority(DataTable dt)
        //{
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        if (dt.Rows[i]["USE_Y"].ToString().Equals("N"))
        //        {
        //            string sFormNm = dt.Rows[i]["PGMID"]?.ToString();
        //            if (string.IsNullOrEmpty(sFormNm))
        //            {
        //                continue;
        //            }

        //            string sBtnNm = dt.Rows[i]["RK"].ToString();

        //            foreach (BarItemLink barItemLink in BarTop.ItemLinks)
        //            {
        //                BarSubItemLink barSubItemLink = barItemLink as BarSubItemLink;
        //                if (barSubItemLink != null)
        //                {
        //                    int iCount = 0;
        //                    foreach (BarItemLink barSubIten in barSubItemLink.Item.ItemLinks)
        //                    {
        //                        string sBarBtnNm = barSubIten.Item.Name;
        //                        if (sBarBtnNm.Equals(sBtnNm))
        //                        {
        //                            (barSubIten.Item as BarButtonItem).Visibility = BarItemVisibility.Never;
        //                            iCount++;
        //                        }
        //                        if (iCount == barSubItemLink.Item.ItemLinks.Count)
        //                        {
        //                            (barSubItemLink.Item as BarSubItem).Visibility = BarItemVisibility.Never;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        private bool FormIsExist(Type tp)
        {
            foreach (Form ff in this.MdiChildren)
            {
                if (ff.GetType() == tp)
                {
                    ff.Focus();
                    ff.BringToFront();
                    ff.AutoScaleMode = AutoScaleMode.Dpi;
                    return true;
                }

            }
            return false;
        }

        private void FormCheck(BarButtonItem barbtn, Form fm)
        {
            
        }

        #region[폼 중복체크]
       
        private void BarbtnItDeptCdDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccDeptCdMgtDev fm = new AccDeptCdMgtDev();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItDealerCdDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccDealerCdDev fm = new AccDealerCdDev();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItAccCdMgtDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC01001F01 fm = new AC01001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItPayAlw_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccSalaryClass fm = new AccSalaryClass();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItPaymentCd_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccPaymentCode fm = new AccPaymentCode();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItAccountNumber_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccAccountNumberMgt fm = new AccAccountNumberMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItCreditCardDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccCreditCardMgt fm = new AccCreditCardMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItCommonCodeMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            CommonCdMgt fm = new CommonCdMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItUnitCostMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            UnitCostMgt fm = new UnitCostMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItJaJae_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            JajaeCdMgt fm = new JajaeCdMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItSlipType_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccAutoSlipType fm = new AccAutoSlipType();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItEquipCdMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EquipCdMgt fm = new EquipCdMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItUserMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            UserMgt fm = new UserMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItConsumeCd_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EquipConsumeCd fm = new EquipConsumeCd();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItPlan_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            ProdPlanMgt fm = new ProdPlanMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        
        private void BarbtnItEquipMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EquipToolMgt fm = new EquipToolMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItConsInOut_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EquipConsumeInOutRpt fm = new EquipConsumeInOutRpt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItClaim_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            QimClaimDev fm = new QimClaimDev();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItQimBadMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            QimBadMgt fm = new QimBadMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItPersonnelRecords_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccPersonnelRecordsDev fm = new AccPersonnelRecordsDev();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItSlip_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC02001F01 fm = new AC02001F01();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItDailyBalance_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC04001F01 fm = new AC04001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();

            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItMeasure_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccInOutMgt fm = new AccInOutMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItClose_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccMeasureCloseDev fm = new AccMeasureCloseDev();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItMeasureDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccMeasureDevImg fm = new AccMeasureDevImg();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItSiteDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccFieldPSDailyRpt fm = new AccFieldPSDailyRpt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItUnitPriceDev_ItemClick(object sender, ItemClickEventArgs e)
        {
            //SplashScreenManager.ShowForm(typeof(WaitForm1));
            //AccInputUnitPrcDev fm = new AccInputUnitPrcDev();
            //if (FormIsExist(fm.GetType()))
            //{
            //    MessageBox.Show("이미 폼이 열려 있습니다");
            //    fm.Dispose();
            //    SplashScreenManager.CloseForm();
            //    return;
            //}
            //else
            //{
            //    fm.MdiParent = this;
            //    fm.Show();
            //}
            //SplashScreenManager.CloseForm();
        }

        private void BarbtnMesMngProgram_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            MesMgtProgramDev fm = new MesMgtProgramDev();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItInComeDailyReport_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            IncomeScrapDailyReport fm = new IncomeScrapDailyReport();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItScrab_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            DailyUnitCost fm = new DailyUnitCost();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItLstSalePlan_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            //MonthlySalePlan fm = new MonthlySalePlan();
            NewSalePlan fm = new NewSalePlan();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItStandardCd_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            StandardCd fm = new StandardCd();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItProdStatus_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            ProdStatus fm = new ProdStatus();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItCompanyInfo_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            CompanyInfo fm = new CompanyInfo();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnDirectDailyRpt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            DailyDirectSend fm = new DailyDirectSend();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        private void BarbtnItDailyMonthlyRpt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            DailyMonthlyForm fm = new DailyMonthlyForm();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EquipInspect fm = new EquipInspect();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        private void BarBtnItCtfctnEmploee_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            HrCertificateOfEmployment fm = new HrCertificateOfEmployment();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnEquipHisMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EquipCdHistoryMgt fm = new EquipCdHistoryMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnItemResetLayout_ItemClick(object sender, ItemClickEventArgs e)
        {
            ResetGridViewLayout fm = new ResetGridViewLayout();
            if (fm.ShowDialog() == DialogResult.OK)
            {
                XtraMessageBox.Show("초기화가 완료되었습니다.");
            }
        }

        private void BarbtnItPurcPfrm_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            TeamPurcPerformance fm = new TeamPurcPerformance();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnItJournal_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC03001F01 fm = new AC03001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC05001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC05001F01 fm = new AC05001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC07001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC07001F01 fm = new AC07001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC06001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC06001F01 fm = new AC06001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BatBtnAC08001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC08001F01 fm = new AC08001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC09001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC09001F01 fm = new AC09001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC10001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC10001F01 fm = new AC10001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC11001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC11001F01 fm = new AC11001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC12001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC18001F01 fm = new AC18001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC13001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC13001F01 fm = new AC13001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarbtnIN07001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            IN07001F01 fm = new IN07001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            ChagmReasonMgt fm = new ChagmReasonMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnSYS001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SYS001F01 fm = new SYS001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC14001F02_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC14001F01 fm = new AC14001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BatBtnIN09001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            //IN09001F01 fm = new IN09001F01();
            IN010F00 fm = new IN010F00();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        private void BarBtnAC16001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC16001F00 fm = new AC16001F00();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BatBtnAC17001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC17001F01 fm = new AC17001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnPD02001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            PD05001F01 fm = new PD05001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }


        private void BatBtnIN10001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            IN10001F01 fm = new IN10001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnSYS002F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SYS002F01 fm = new SYS002F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BatBtnSYS003F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SYS003F01 fm = new SYS003F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC19001F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC19001F00 fm = new AC19001F00();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAC20001F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AC20001F00 fm = new AC20001F00();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BatBtnIN11001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            IN11001F01 fm = new IN11001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        private void BarBtnPD04001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            PD04001F01 fm = new PD04001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BatBtnIN12001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            IN12001F01 fm = new IN12001F01();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        //DB마이그 작업을 위한 임시메뉴
        private void BarBtnSYS010F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SYS010F00 fm = new SYS010F00();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.Owner = this;
                fm.ShowDialog();
            }
        }

        private void BarbtnItLstSaleResult_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            PurTeamRecord fm = new PurTeamRecord();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnBusinessReport_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            BusinessReport fm = new BusinessReport();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnRptApplSystem_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            RptApplSystem fm = new RptApplSystem();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnAllReport_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AllReport fm = new AllReport();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnDoctMgt_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            DoctMgt fm = new DoctMgt();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnGE001F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            GE001F00 fm = new GE001F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnGE002F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            GE002F00 fm = new GE002F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnGE003F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            GE003F00 fm = new GE003F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnSL001F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SL001F00 fm = new SL001F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnSL001F01_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SL001F01 fm = new SL001F01();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnIN001F02_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            IN001F02 fm = new IN001F02();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnSL002F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SL002F00 fm = new SL002F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        private void BarBtnSC010F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            SC010F00 fm = new SC010F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void BarBtnEQ001F00_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            EQ001F00 fm = new EQ001F00();

            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
        #endregion
        
        private void FmMainToolBar2_MdiChildActivate(object sender, EventArgs e)
        {
            BarStaticTxtFormText.Caption = "";
            if (this.ActiveMdiChild != null)
                BarStaticCurPage.Caption = this.ActiveMdiChild.Name;
        }
        private void FmMainToolBar_FormClosing(object sender, FormClosingEventArgs e)
        {
            FmLogin fm = new FmLogin();
            fm.Close();
            Application.Exit();
        }
        private void FmMainToolBar2_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                //_T_CUR_TIME.Abort();
                //_T_ALARM.Abort();
                //_T_ALARMIMG.Abort();
                //_T_F_ALARM.Abort();

                DBConn.DbDisConn(DBConn.dbCon);

                //전자결재 임시데이터 삭제
                string tempFolder = Application.StartupPath + "/docttemp/" + drUser["USRCD"] + "/";
                DirectoryInfo di = new DirectoryInfo(tempFolder);

                if (File.Exists(tempFolder))
                    di.Delete(true);

                if (System.Windows.Forms.Application.MessageLoop)
                {
                    // WinForms app
                    System.Windows.Forms.Application.Exit();
                }
                else
                {
                    // Console app
                    System.Environment.Exit(1);
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        
        /*
         * 현재 폼에서 자식폼의 Text를 바꿈
         * 자식폼에서 TextChanged 이벤트에서 Save처리
        */
        private void BatBtnSave_ItemClick(object sender, ItemClickEventArgs e)
        {
            string sFormName = BarManagerFmMain.ActiveMdiChild?.Name;
            if (string.IsNullOrEmpty(sFormName))
                return;

            if (XtraMessageBox.Show(BarManagerFmMain.ActiveMdiChild.Text + "의 Layout정보를 저장하시겠습니까?", "Layout저장여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            Form frm = BarManagerFmMain.ActiveMdiChild;
            ComnEtcFunc.SetFormText(frm);
        }
        
        private void BarBtnSaveLayout_ItemClick(object sender, ItemClickEventArgs e)
        {
            BatBtnSave_ItemClick(null, null);
        }

        private void BatBtnItemInit_ItemClick(object sender, ItemClickEventArgs e)
        {
            string sPgmId = BarManagerFmMain.ActiveMdiChild?.Name;
            if (string.IsNullOrEmpty(sPgmId))
            {
                return;
            }

            if (XtraMessageBox.Show("선택한 프로그램의 레이아웃 설정값이 사라지게 됩니다. \r\n그래도 초기화 하시겠습니까?"
               , "레이아웃 초기화 여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            string sId = UserID;

            if (System.IO.Directory.Exists(Application.StartupPath + @"\xaml\" + sId))
            {
                string[] files = System.IO.Directory.GetFiles(Application.StartupPath + @"\xaml\" + sId);

                foreach (string s in files)
                {
                    Cursor = Cursors.WaitCursor;
                    string fileName = System.IO.Path.GetFileName(s);
                    if (fileName.Contains(sPgmId))
                    {
                        string deletefile = Application.StartupPath + @"\xaml\" + sId + @"\" + fileName;
                        System.IO.File.Delete(deletefile);

                        Cursor = Cursors.Default;
                    }
                    Cursor = Cursors.Default;
                }
            }

            XtraMessageBox.Show("초기화가 완료되었습니다.\r\n해당 폼이 다시 실행됩니다.");

            string sFormName = BarManagerFmMain.ActiveMdiChild.Name;
            BarManagerFmMain.ActiveMdiChild.Dispose();
            XtraForm frm = (XtraForm)GetAssemblyForm(sFormName);
            
            SplashScreenManager.ShowForm(typeof(WaitForm1));

            if (FormIsExist(frm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                frm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                frm.MdiParent = this;
                frm.Show();
            }
            SplashScreenManager.CloseForm();

            Cursor = Cursors.Default;
        }
        
        public static Form GetAssemblyForm(string strFormName)
        {
            Form f = null;
            foreach (Type t in System.Reflection.Assembly.GetExecutingAssembly().GetTypes())
            {
                if (t.Name == strFormName)                     //프로젝트 내 폼 중에서 찾을 이름과 같으면...
                {
                    object o = Activator.CreateInstance(t);    //인스턴스 개체 생성 
                    f = o as Form;                                  //인스턴스 개체 폼 형식으로 캐스팅
                }
            }
            return f;
        }

        private void BarBtnCancelAll_ItemClick(object sender, ItemClickEventArgs e)
        {
            foreach (Form child in this.MdiChildren)
            {
                child.Dispose();
            }
        }

        private void BarBtnSelectSkin_ItemClick(object sender, ItemClickEventArgs e)
        {
            SYS004F01 form = new SYS004F01();
            form.Show();
        }

        

        private void BarBtnAlarm_ItemClick(object sender, ItemClickEventArgs e)
        {
            int height = BarTop.FloatSize.Height;

            Form fc = Application.OpenForms["PopAlarm"];

            if(fc != null)
            {
                fc.Close();
            }
            else
            {
                PopAlarm frm = new PopAlarm();
                frm.Owner = this;
                frm.iHeight = height;
                frm.Show();
            }
        }

        

        #region 알람 창
        private delegate void OpenAlarmDelegate(string msg);
        private OpenAlarmDelegate openAlarmDelegate = null;

        private void Alarm()
        {
            StringBuilder strSql = new StringBuilder();

            while (true)
            {
                try
                {
                    strSql.Clear();
                    strSql.AppendLine(" SELECT COUNT(*) AS CNT");
                    strSql.AppendLine("   FROM ALAMMGT  ");
                    strSql.AppendLine("  WHERE CDATE > '"+ LoginTime + "'");
                    strSql.AppendLine("    AND USRCD = " + drUser["USRCD"]);

                    DataTable dt = DBConn.GetDataTable(strSql.ToString());

                    if(dt != null)
                    {
                        double dCnt = double.Parse(dt.Rows[0]["CNT"]?.ToString());

                        if(dCnt > 0)
                        {
                            Invoke(openAlarmDelegate, "새 알림이 도착했습니다!");
                            LoginTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                            Thread.Sleep(3000);
                        }
                    }
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }
            }

        }

        private void GetAlarm()
        {
            if (_T_ALARM == null)
            {
                _T_ALARM = new Thread(new ThreadStart(Alarm));
            }

            if (_T_ALARM.IsAlive)
            {
                _T_ALARM.Abort();
            }

            _T_ALARM.IsBackground = true;
            _T_ALARM.Start();
        }

        private void openAlarmForm(string msg)
        {
            int height = BarTop.FloatSize.Height;

            Form fc2 = Application.OpenForms["PopAlarm"];

            if (fc2 != null)
            {
                fc2.Close();
            }

            PopAlarm frm2 = new PopAlarm();
            frm2.Owner = this;
            frm2.iHeight = height;
            frm2.Show();

            //Form fc = Application.OpenForms["PopAlarmMsg"];

            //if (fc != null)
            //{
            //    fc.Close();
            //}

            //PopAlarmMsg frm = new PopAlarmMsg();
            //frm.Owner = this;
            //frm._MSG = msg;
            //frm.Show();
        }
        #endregion

        #region 알람 종 이미지
        private delegate void ChangeAlarmImgDelegate(double iCnt);
        private ChangeAlarmImgDelegate changeAlarmImgDelegate = null;

        private void SetAlarmImg()
        {
            if (_T_ALARMIMG == null)
            {
                _T_ALARMIMG = new Thread(new ThreadStart(setAlarmImgFun));
            }

            if (_T_ALARMIMG.IsAlive)
            {
                _T_ALARMIMG.Abort();
            }

            _T_ALARMIMG.IsBackground = true;
            _T_ALARMIMG.Start();
        }

        private void setAlarmImgFun()
        {
            StringBuilder strSql = new StringBuilder();

            while (true)
            {
                try
                {
                    strSql.Clear();
                    strSql.AppendLine(" SELECT COUNT(*) AS CNT");
                    strSql.AppendLine("   FROM ALAMMGT  ");
                    strSql.AppendLine("  WHERE READYN = 'N'");
                    strSql.AppendLine("    AND USRCD = " + drUser["USRCD"]);

                    DataTable dt = DBConn.GetDataTable(strSql.ToString());

                    if (dt != null)
                    {
                        double dCnt = double.Parse(dt.Rows[0]["CNT"]?.ToString());

                        Invoke(changeAlarmImgDelegate, dCnt);
                    }

                    Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }

        private void ChangeAlarmImg(double iCnt)
        {
            if (iCnt == 0)
            {
                BarBtnAlarm.ImageOptions.Image = Properties.Resources.bell1;
                BarBtnAlarm.Caption = "새 알림 없음";
            }
            else if (iCnt > 0)
            {
                BarBtnAlarm.ImageOptions.Image = Properties.Resources.notification1;
                BarBtnAlarm.Caption = iCnt + "개의 새 알림";
            }
        }
        #endregion

        #region 로그인 후 첫 알림창
        private void SetFAlarm()
        {
            if (_T_F_ALARM == null)
            {
                _T_F_ALARM = new Thread(new ThreadStart(fstLoginAlarm));
            }

            if (_T_F_ALARM.IsAlive)
            {
                _T_F_ALARM.Abort();
            }

            _T_F_ALARM.IsBackground = true;
            _T_F_ALARM.Start();
        }

        private void fstLoginAlarm()
        {
            try
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT COUNT(*) AS CNT");
                strSql.AppendLine("   FROM ALAMMGT  ");
                strSql.AppendLine("  WHERE READYN = 'N'");
                strSql.AppendLine("    AND USRCD = " + drUser["USRCD"]);

                DataTable dt = DBConn.GetDataTable(strSql.ToString());

                if (dt != null)
                {
                    double dCnt = double.Parse(dt.Rows[0]["CNT"]?.ToString());

                    if (dCnt > 0)
                    {
                        BarBtnAlarm.ImageOptions.Image = Properties.Resources.notification1;
                        BarBtnAlarm.Caption = dCnt + "개의 새 알림";

                        Invoke(openAlarmDelegate, dCnt+"개의 새 알림이 있습니다.");
                    }
                    else
                    {
                        BarBtnAlarm.ImageOptions.Image = Properties.Resources.bell1;
                        BarBtnAlarm.Caption = "새 알림 없음";
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 메인 폼 선택 시 알림창 닫기
        private void FmMainToolBar2_Activated(object sender, EventArgs e)
        {
            Form fc = Application.OpenForms["PopAlarm"];

            if (fc != null)
            {
                fc.Close();
            }
        }


        #endregion

        private void barBtnAccFieldCustom_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccFieldCustom fm = new AccFieldCustom();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void barBtnAccFidelVs_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccFieldVs fm = new AccFieldVs();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void barBtnItPlanA_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            ProdPlanA fm = new ProdPlanA();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }

        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            AccMeasureDevImg fm = new AccMeasureDevImg();
            if (FormIsExist(fm.GetType()))
            {
                MessageBox.Show("이미 폼이 열려 있습니다.");
                fm.Dispose();
                SplashScreenManager.CloseForm();
                return;
            }
            else
            {
                fm.MdiParent = this;
                fm.Show();
            }
            SplashScreenManager.CloseForm();
        }
    }
}