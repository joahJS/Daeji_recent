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
using OpenCvSharp;
using System.Threading;
using OpenCvSharp.Extensions;

namespace Daeji_MONITERING
{
    public partial class CV001F00 : DevExpress.XtraEditors.XtraForm
    {
        public CV001F00()
        {
            InitializeComponent();
        }

        #region [ CCTV 영상 변수 ]
        Mat frame;
        public string _RSTPaddr = string.Empty;
        #endregion
        #region [ 쓰레드 ]
        //cctv
        Thread t_Camera;
        #endregion
        #region [ 녹화 및 저장 변수 ]
        VideoCapture capture;
        #endregion
        #region [ 조건 변수 ]
        bool isFirstCamera;
        bool isConnected;
        #region [ CCTV 연결 시작/종료 구분용 변수 ]
        private bool _isWaitStart = true;
        #endregion
        #endregion

        private volatile bool _shouldStop;

        private void CV001F00_Load(object sender, EventArgs e)
        {

        }

        private void CV001F00_Shown(object sender, EventArgs e)
        {
            StartAll();
        }

        // 쓰레드 시작 메서드
        private void StartAll()
        {
            if (t_Camera == null || !t_Camera.IsAlive)
            {
                t_Camera = new Thread(new ThreadStart(StartCamera));
                t_Camera.IsBackground = true;
                t_Camera.Start();
            }
        }

        //카메라 쓰레드 (반복)
        private void StartCamera()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        if (!isConnected)
                        {
                            string RSTPaddr = _RSTPaddr;
                            capture = new VideoCapture(RSTPaddr);
                            frame = new Mat();
                            isConnected = true;
                        }

                        if (!isFirstCamera)
                        {
                            isFirstCamera = true;
                        }

                        if (!capture.Read(frame))
                        {
                            Cv2.WaitKey();
                        }

                        Bitmap bitmap2 = BitmapConverter.ToBitmap(frame);
                        PicCCTV.Image = bitmap2;

                        if (_isWaitStart)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 연결");
                            _isWaitStart = false;
                        }

                        if (Cv2.WaitKey(1) >= 0)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                        if (!_isWaitStart)
                        {
                            //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 단절");
                            _isWaitStart = true;
                        }
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}