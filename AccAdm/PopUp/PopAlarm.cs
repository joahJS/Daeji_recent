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
using System.Timers;
using ComLib;
using DevExpress.Utils;
using System.Data.SqlClient;
using DevExpress.XtraSplashScreen;

namespace AccAdm
{
    public partial class PopAlarm : DevExpress.XtraEditors.XtraForm
    {
        #region 높이/위쪽 위치 설정하기 대리자 - SetHeightTopDelegate(flag)

        /// <summary>
        /// 높이/위쪽 위치 설정하기 대리자
        /// </summary>
        /// <param name="flag">플래그</param>
        private delegate void SetHeightTopDelegate(int flag);

        #endregion

        #region 넓이/왼쪽 위치 설정하기 대리자 - SetWidthRightDelegate(flag) 정은영추가

        /// <summary>
        /// 넓이/왼쪽 위치 설정하기 대리자
        /// </summary>
        /// <param name="flag">플래그</param>
        private delegate void SetWidthRightDelegate(int flag);

        #endregion

        #region Field

        /// <summary>
        /// 높이/위쪽 위치 설정하기 대리자
        /// </summary>
        private SetHeightTopDelegate setHeightTopDelegate = null;

        /// <summary>
        /// 넓이/왼쪽 위치 설정하기 대리자 정은영추가
        /// </summary>
        private SetWidthRightDelegate setWidthRightDelegate = null;

        /// <summary>
        /// 타이머
        /// </summary>
        private System.Timers.Timer timer;

        //부모폼 위치
        private Point parentLocation;

        //부모폼 넓이
        private int parentWidth;
        //부모폼 높이
        private int parentHeight;


        //
        public int iHeight;

        #endregion

        #region 생성자 - NoticeForm()
        /// <summary>
        /// 생성자
        /// </summary>
        public PopAlarm()
        {
            InitializeComponent();
        }
        #endregion

        #region 폼 로드시 처리하기 - Form_Load(sender, e)
        /// <summary>
        /// 폼 로드시 처리하기
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void Form_Load(object sender, EventArgs e)
        {
            FmMainToolBar2 pFrm = (FmMainToolBar2) this.Owner;
            parentLocation = pFrm.Location;
            parentWidth = pFrm.Width;
            parentHeight = pFrm.Height;

            //setHeightTopDelegate = new SetHeightTopDelegate(SetHeightTop);
            setWidthRightDelegate = new SetWidthRightDelegate(SetWidthRight);

            Size = new Size(0, parentHeight - int.Parse(Math.Round(parentHeight*0.062,0).ToString()) - iHeight);

            Location = new Point(parentLocation.X + parentWidth - Width, parentLocation.Y + parentHeight - Height - int.Parse(Math.Round(parentHeight*0.035,0).ToString()));
            //Location = new Point(parentLocation.X + parentWidth - Width , parentLocation.Y + parentHeight - Height- 37);

            this.timer = new System.Timers.Timer(2);
            this.timer.Elapsed += timer_Elapsed_PopUp2;
            
            this.timer.Start();
            
            GetAlarmData();
            DelAlarmData();


            //e.Item.AppearanceItem.Normal.BackColor = Color.FromArgb(64, 64, 64);

        }
        #endregion

        #region 타이머 경과시 처리하기 (팝업용) - timer_Elapsed_PopUp(sender, e)
        /// <summary>
        /// 타이머 경과시 처리하기 (팝업용)
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void timer_Elapsed_PopUp(object sender, ElapsedEventArgs e)
        {
            if (Height < 140)
            {
                Invoke(setHeightTopDelegate, 0);
            }
            else
            {
                this.timer.Stop();

                this.timer.Elapsed -= timer_Elapsed_PopUp;
                this.timer.Elapsed += timer_Elapsed_PopOut;

                this.timer.Interval = 3000;

                this.timer.Start();
            }

            Application.DoEvents();
        }
        #endregion

        #region 타이머 경과시 처리하기 (팝업용) - timer_Elapsed_PopUp2(sender, e) 정은영추가
        /// <summary>
        /// 타이머 경과시 처리하기 (팝업용)
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void timer_Elapsed_PopUp2(object sender, ElapsedEventArgs e)
        {
            if (Width < 400)
            {
                Invoke(setWidthRightDelegate, 0);
            }
            else
            {
                this.timer.Stop();
                //UpdateUIControl();
                //this.timer.Elapsed -= timer_Elapsed_PopUp2;
                //this.timer.Elapsed += timer_Elapsed_PopOut2;

                //this.timer.Interval = 3000;

                //this.timer.Start();
            }
            //Application.DoEvents();
        }
        //private void UpdateUIControl()
        //{
        //    // UI 컨트롤의 업데이트 작업 수행
        //    if (this.InvokeRequired)
        //    {
        //        this.Invoke(new Action(UpdateUIControl));
        //    }
        //    else
        //    {
        //        // UI 컨트롤의 업데이트 작업
        //        this.Text = "새로운 값"; // 예시: UI 컨트롤의 Text 속성 업데이트
        //    }
        //}

        #endregion

        #region 타이머 경과시 처리하기 (팝아웃용) - timer_Elapsed_PopOut(sender, e)
        /// <summary>
        /// 타이머 경과시 처리하기 (팝아웃용)
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void timer_Elapsed_PopOut(object sender, ElapsedEventArgs e)
        {
            while (Height > 2)
            {
                Invoke(setHeightTopDelegate, 1);
            }

            this.timer.Stop();

            Application.DoEvents();

            Invoke(setHeightTopDelegate, 2);
        }
        #endregion

        #region 타이머 경과시 처리하기 (팝아웃용) - timer_Elapsed_PopOut2(sender, e) 정은영 추가
        /// <summary>
        /// 타이머 경과시 처리하기 (팝아웃용)
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void timer_Elapsed_PopOut2(object sender, ElapsedEventArgs e)
        {
            while (Width > 2)
            {
                Invoke(setWidthRightDelegate, 1);
            }

            this.timer.Stop();

            Application.DoEvents();

            Invoke(setWidthRightDelegate, 2);
        }
        #endregion

        #region 높이/위쪽 위치 설정하기 - SetHeightTop(flag)
        /// <summary>
        /// 높이/위쪽 위치 설정하기
        /// </summary>
        /// <param name="flag">플래그</param>
        private void SetHeightTop(int flag)
        {
            if (flag == 0)
            {
                Height++;

                Top--;
            }
            else if (flag == 1)
            {
                Height--;

                Top++;
            }
            else if (flag == 2)
            {
                Close();
            }
        }
        #endregion

        #region 넓이/좌측 위치 설정하기 - SetWidthLeft(flag) - 정은영 추가
        /// <summary>
        /// 높이/위쪽 위치 설정하기
        /// </summary>
        /// <param name="flag">플래그</param>
        private void SetWidthRight(int flag)
        {
            if (flag == 0)
            {
                //Width++;
                //Left--;

                Width += 40;
                Left -= 40;
            }
            else if (flag == 1)
            {
                //Width--;
                //Left++;

                Width -= 40;
                Left += 40;
            }
            else if (flag == 2)
            {
                Close();
            }
        }
        #endregion

        private DataTable GetAlarmData()
        {
            DataTable dt = new DataTable();

            try
            {
                string usrcd = FmMainToolBar2.drUser["USRCD"]?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT ALRMCD     ");
                strSql.AppendLine("      , USRCD      ");
                strSql.AppendLine("      , PGMID      ");
                strSql.AppendLine("      , PGMNM      ");
                strSql.AppendLine("      , ALMSG      ");
                strSql.AppendLine("      , ALKEY1     ");
                strSql.AppendLine("      , ALKEY2     ");
                strSql.AppendLine("      , ALKEY3     ");
                strSql.AppendLine("      , CDATE      ");
                strSql.AppendLine("   FROM ALAMMGT    ");
                strSql.AppendLine("  WHERE USRCD = '"+ usrcd + "'");
                strSql.AppendLine("    AND READYN = 'N'");
                strSql.AppendLine("  ORDER BY CDATE DESC ");

                dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                GridRetr.DataSource = dt;

                
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

            return dt;
        }

        private void GridViewTile_ItemCustomize(object sender, DevExpress.XtraGrid.Views.Tile.TileViewItemCustomizeEventArgs e)
        {
            e.Item.AppearanceItem.Normal.BackColor = Color.FromArgb(64, 64, 64);
        }

        private void GridViewTile_ContextButtonClick(object sender, ContextItemClickEventArgs e)
        {
            if(e.Item.Name == "BtnReadY")
            {
                try
                {
                    StringBuilder strSql = new StringBuilder();
                    string sALRMCD = GridViewTile.GetFocusedRowCellValue("ALRMCD")?.ToString();

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE ALAMMGT      ");
                    strSql.AppendLine("    SET READYN = 'Y' ");
                    strSql.AppendLine("  WHERE ALRMCD = "+ sALRMCD + "   ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;

                    DataTable dt = GetAlarmData();

                    if(dt == null || dt.Rows.Count == 0)
                    {
                        Close();
                    }
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }

        private void GridViewTile_ItemClick(object sender, DevExpress.XtraGrid.Views.Tile.TileViewItemClickEventArgs e)
        {
            string sALRMCD = GridViewTile.GetFocusedRowCellValue("ALRMCD")?.ToString();
            string sUSRCD  = GridViewTile.GetFocusedRowCellValue("USRCD")?.ToString();
            string sPGMID  = GridViewTile.GetFocusedRowCellValue("PGMID")?.ToString();
            string sPGMNM  = GridViewTile.GetFocusedRowCellValue("PGMNM")?.ToString();
            string sALKEY1 = GridViewTile.GetFocusedRowCellValue("ALKEY1")?.ToString();

            StringBuilder strSql = new StringBuilder();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" UPDATE ALAMMGT      ");
                strSql.AppendLine("    SET READYN = 'Y' ");
                strSql.AppendLine("  WHERE ALRMCD = " + sALRMCD + "   ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                GetAlarmData();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }

            if (sPGMID == "RptApplSystem") //전자결재
            {
                Form fc = Application.OpenForms["RptApplSystem"];

                if (fc != null)
                {
                    fc.Close();
                }

                RptApplSystem frm = new RptApplSystem();
                frm._FINDOBJ = 0;
                frm._FINDWORD = sALKEY1;
                frm.MdiParent = Application.OpenForms["FmMainToolBar2"];
                frm.Show();
                frm.SearchFun();

                Close();
            }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dtn = new DataTable();
            string usrcd = FmMainToolBar2.drUser["USRCD"]?.ToString();
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT ALRMCD     ");
            strSql.AppendLine("      , USRCD      ");
            strSql.AppendLine("      , PGMID      ");
            strSql.AppendLine("      , PGMNM      ");
            strSql.AppendLine("      , ALMSG      ");
            strSql.AppendLine("      , ALKEY1     ");
            strSql.AppendLine("      , ALKEY2     ");
            strSql.AppendLine("      , ALKEY3     ");
            strSql.AppendLine("      , CDATE      ");
            strSql.AppendLine("   FROM ALAMMGT    ");
            strSql.AppendLine("  WHERE USRCD = '" + usrcd + "'");
            strSql.AppendLine("    AND READYN = 'N'");
            strSql.AppendLine("  ORDER BY CDATE DESC ");

            dtn = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dtn.Rows.Count; i++)
            {
                try
                {
                    string sALRMCD = dtn.Rows[i]["ALRMCD"]?.ToString();

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE ALAMMGT      ");
                    strSql.AppendLine("    SET READYN = 'Y' ");
                    strSql.AppendLine("  WHERE ALRMCD = " + sALRMCD + "   ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;

                    GetAlarmData();
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }
        private void DelAlarmData()
        {
            DataTable dtn = new DataTable();
            string usrcd = FmMainToolBar2.drUser["USRCD"]?.ToString();
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT ALRMCD     ");
            strSql.AppendLine("      , USRCD      ");
            strSql.AppendLine("      , PGMID      ");
            strSql.AppendLine("      , PGMNM      ");
            strSql.AppendLine("      , ALMSG      ");
            strSql.AppendLine("      , ALKEY1     ");
            strSql.AppendLine("      , ALKEY2     ");
            strSql.AppendLine("      , ALKEY3     ");
            strSql.AppendLine("      , CDATE      ");
            strSql.AppendLine("   FROM ALAMMGT    ");
            strSql.AppendLine("  WHERE USRCD = '" + usrcd + "'");
            strSql.AppendLine("    AND READYN = 'N'");
            strSql.AppendLine("  ORDER BY CDATE DESC ");

            dtn = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dtn.Rows.Count; i++)
            {
                try
                {
                    string sALRMCD = dtn.Rows[i]["ALRMCD"]?.ToString();

                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE ALAMMGT      ");
                    strSql.AppendLine("    SET READYN = 'Y' ");
                    strSql.AppendLine("  WHERE ALRMCD = " + sALRMCD + "   ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }
    }
}