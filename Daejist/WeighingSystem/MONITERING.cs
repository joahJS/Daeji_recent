using DevExpress.XtraEditors;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Size =  System.Drawing.Size;

namespace WeighingSystem
{
    public partial class MONITERING : Form
    {
        
        #region [ CCTV 영상 변수 ]
        Mat frame1;
        Mat frame2;
        Mat frame3;
        string RSTPaddr1 = string.Empty;
        string RSTPaddr2 = string.Empty;
        string RSTPaddr3 = string.Empty;
        #endregion
        #region [ 쓰레드 ]
        Thread t_Camera1;
        Thread t_Camera2;
        Thread t_Camera3;
        #endregion
        #region [ 녹화 및 저장 변수 ]
        VideoCapture capture1;
        VideoCapture capture2;
        VideoCapture capture3;
        #endregion
        #region [ 조건 변수 ]
        bool isFirstCamera;
        bool isConnected;
        #region [ CCTV 연결 시작/종료 구분용 변수 ]
        private bool _isWaitStart_1 = true;
        private bool _isWaitStart_2 = true;
        private bool _isWaitStart_3 = true;
        #endregion
        #endregion

        private volatile bool _shouldStop;

        public MONITERING()
        {
            InitializeComponent();
        }

        private void MONITERING_Shown(object sender, EventArgs e)
        {
            StartAll();
        }

        private void MONITERING_Load(object sender, EventArgs e)
        {
            RTSP_Init();
        }

        // 카메라 RTSP 정보 불러오기 메서드
        private void RTSP_Init()
        {
            
            //RSTPaddr2 = "rtsp://fa356df5:caps%40112@10.255.254.3/H.264";
            RSTPaddr1 = "rtsp://fa356df5:caps%40112@192.168.1.241/h264";
            RSTPaddr2 = "rtsp://fa356df5:caps%40112@192.168.1.242/h264";
            RSTPaddr3 = "rtsp://fa356df5:caps%40112@192.168.1.244/h264";
            //RSTPaddr2 = "rtsp://caps1:caps%40112@10.255.254.3//h264";

        }

        #region [CCTV 영상 가져오기 쓰레드]
        // 쓰레드 시작 메서드
        private void StartAll()
        {
            if (t_Camera1 == null || !t_Camera1.IsAlive)
            {
                t_Camera1 = new Thread(new ThreadStart(StartCamera1));
                t_Camera1.IsBackground = true;
                t_Camera1.Start();
            }
            if (t_Camera2 == null || !t_Camera2.IsAlive)
            {
                t_Camera2 = new Thread(new ThreadStart(StartCamera2));
                t_Camera2.IsBackground = true;
                t_Camera2.Start();
            }
            if (t_Camera3 == null || !t_Camera3.IsAlive)
            {
                t_Camera3 = new Thread(new ThreadStart(StartCamera3));
                t_Camera3.IsBackground = true;
                t_Camera3.Start();
            }
        }

        // 1번 카메라 쓰레드 (반복)
        private void StartCamera1()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        if (!isConnected)
                        {
                            string RSTPaddr = RSTPaddr1;
                            capture1 = new VideoCapture(RSTPaddr);
                            frame1 = new Mat();
                            isConnected = true;
                        }

                        if (!isFirstCamera)
                        {
                            isFirstCamera = true;
                        }

                        if (!capture1.Read(frame1))
                        {
                            Cv2.WaitKey();
                        }

                        Bitmap bitmap2 = BitmapConverter.ToBitmap(frame1);
                        pictureBox1.Image = bitmap2;

                        if (_isWaitStart_1)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 연결");
                            _isWaitStart_1 = false;
                        }

                        if (Cv2.WaitKey(1) >= 0)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                        if (!_isWaitStart_1)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 단절");
                            _isWaitStart_1 = true;
                        }
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }

        // 2번 카메라 쓰레드 (반복)
        private void StartCamera2()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        if (!isConnected)
                        {
                            string RSTPaddr = RSTPaddr2;
                            capture2 = new VideoCapture(RSTPaddr);
                            frame2 = new Mat();
                            isConnected = true;
                        }

                        if (!isFirstCamera)
                        {
                            isFirstCamera = true;
                        }

                        if (!capture2.Read(frame2))
                        {
                            Cv2.WaitKey();
                        }

                        Bitmap bitmap2 = BitmapConverter.ToBitmap(frame2);
                        pictureBox2.Image = bitmap2;

                        if (_isWaitStart_2)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 연결");
                            _isWaitStart_2 = false;
                        }

                        if (Cv2.WaitKey(1) >= 0)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                        if (!_isWaitStart_2)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 단절");
                            _isWaitStart_2 = true;
                        }
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }
        private void StartCamera3()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        if (!isConnected)
                        {
                            string RSTPaddr = RSTPaddr3;
                            capture3 = new VideoCapture(RSTPaddr);
                            frame3 = new Mat();
                            isConnected = true;
                        }

                        if (!isFirstCamera)
                        {
                            isFirstCamera = true;
                        }

                        if (!capture3.Read(frame3))
                        {
                            Cv2.WaitKey();
                        }

                        Bitmap bitmap3 = BitmapConverter.ToBitmap(frame3);
                        pictureBox3.Image = bitmap3;

                        if (_isWaitStart_3)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "3번 카메라 연결");
                            _isWaitStart_3 = false;
                        }

                        if (Cv2.WaitKey(1) >= 0)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                        if (!_isWaitStart_3)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 단절");
                            _isWaitStart_3 = true;
                        }
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }
        #endregion

        #region [크로스 스레드]
        delegate void CrossThreadSafetySetText(Control ctl, object obj);
        private void CSafeSetText(Control ctl, object obj)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetText(CSafeSetText), ctl, obj);
            else
                ctl.Text = obj.ToString();
        }

        delegate void CrossThreadSafetySetForeColor(Control ctl, Color color);
        private void CSafeSetForeColor(Control ctl, Color color)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetForeColor(CSafeSetForeColor), ctl, color);
            else
                ctl.ForeColor = color;
        }

        delegate void CrossThreadSafetySetBackgroundImage(Control ctl, Image img);
        private void CSafeSetBackgroundImage(Control ctl, Image img)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetBackgroundImage(CSafeSetBackgroundImage), ctl, img);
            else
                ctl.BackgroundImage = img;
        }

        delegate void CrossThreadSafetySetLocation(Control ctl, System.Drawing.Point pt);
        private void CSafeSetLocation(Control ctl, System.Drawing.Point pt)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetLocation(CSafeSetLocation), ctl, pt);
            else
                ctl.Location = pt;
        }

        delegate void CrossThreadSafetySetImage(PictureEdit ctl, Image img);
        private void CSafeSetImage(PictureEdit ctl, Image img)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetImage(CSafeSetImage), ctl, img);
            else
                ctl.Image = img;
        }

        delegate void CrossThreadSafetySetImageBox(PictureBox ctl, Image img);
        private void CSafeSetImageBox(PictureBox ctl, Image img)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetImageBox(CSafeSetImageBox), ctl, img);
            else
                ctl.Image = img;
        }
        delegate void CrossThreadSafetySetTextMemo(MemoEdit ctl, string str);
        private void CSafeSetmemoEdit(MemoEdit ctl, string str)
        {
            if (ctl.InvokeRequired)
            {
                ctl.Invoke(new CrossThreadSafetySetTextMemo(CSafeSetmemoEdit), ctl, str);
            }
            else
            {
                ctl.Text = str;
            }
        }
        #endregion

        #region [프로그램 종료]
        private void label1_Click(object sender, EventArgs e)
        {
            _shouldStop = true;
            Dispose();
        }
        #endregion

        #region [버튼이벤트]
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        

        public void btnCapture_Click(object sender, EventArgs e)
        {

            Bitmap CpBitmap = BitmapConverter.ToBitmap(frame1);
            CpBitmap.Save("D:/test.jpg", ImageFormat.Jpeg);

        }

        #endregion



    }
}

