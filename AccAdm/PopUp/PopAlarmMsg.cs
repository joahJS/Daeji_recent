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
using System.Runtime.InteropServices;

namespace AccAdm
{
    public partial class PopAlarmMsg : DevExpress.XtraEditors.XtraForm
    {
        #region 높이/위쪽 위치 설정하기 대리자 - SetHeightTopDelegate(flag)

        /// <summary>
        /// 높이/위쪽 위치 설정하기 대리자
        /// </summary>
        /// <param name="flag">플래그</param>
        private delegate void SetHeightTopDelegate(int flag);

        #endregion

        #region Field

        /// <summary>
        /// 높이/위쪽 위치 설정하기 대리자
        /// </summary>
        private SetHeightTopDelegate setHeightTopDelegate = null;

        /// <summary>
        /// 타이머
        /// </summary>
        private System.Timers.Timer timer;


        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn( int nLeftRect
                                                       , int nTopRect
                                                       , int nRightRect
                                                       , int nBottomRect
                                                       , int nWidthEllipse
                                                       , int nHeightEllipse);

        public string _MSG = string.Empty;

        #endregion


        public PopAlarmMsg()
        {
            InitializeComponent();
        }

        private void PopAlarmMsg_Load(object sender, EventArgs e)
        {
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, this.Width, this.Height, 20, 20));
            label1.Text = _MSG;

            setHeightTopDelegate = new SetHeightTopDelegate(SetHeightTop);

            Size = new Size(200, 0);
            Location = new Point(Screen.PrimaryScreen.WorkingArea.Width - Width - 20, Screen.PrimaryScreen.WorkingArea.Height-1);

            this.timer = new System.Timers.Timer(2);

            this.timer.Elapsed += timer_Elapsed_PopUp;

            this.timer.Start();
        }

        #region 타이머 경과시 처리하기 (팝업용) - timer_Elapsed_PopUp(sender, e)
        /// <summary>
        /// 타이머 경과시 처리하기 (팝업용)
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void timer_Elapsed_PopUp(object sender, ElapsedEventArgs e)
        {
            if (Height < 100)
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

        #region 높이/위쪽 위치 설정하기 - SetHeightTop(flag)
        /// <summary>
        /// 높이/위쪽 위치 설정하기
        /// </summary>
        /// <param name="flag">플래그</param>
        private void SetHeightTop(int flag)
        {
            if (flag == 0)
            {
                Height+=5;

                Top-=5;
            }
            else if (flag == 1)
            {
                Height-=5;

                Top+=5;
            }
            else if (flag == 2)
            {
                Close();
            }
        }
        #endregion

        private void pictureEdit2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}