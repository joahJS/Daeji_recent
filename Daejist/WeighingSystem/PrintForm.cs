using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
/*
 * 작성일자 
 * 작성자 : 앤로비
 * ----------------HISTORY------------------
 * 1. 수정일자 : 2021-01-28
 *    수정자 : 고혜성
 *    수정내용 : (현업요청)메시지 폼이 나온 후 10초간 아무런 액션이 없을 시 취소처리하도록 수정
 *    
 *    
 * 2. 수정일자 : 2022-12-08
 *    수정자   : 정은영
 *    수정내용 : (요청)
 *              1. (기존)완료버튼 눌렀을때 캡쳐프로그램 닫기 => (변경)계근표 출력 이후 닫기 진행
 */

namespace WeighingSystem
{
    public partial class PrintForm : Form
    {
        public PrintForm()
        {
            InitializeComponent();
        }

        private void buttonYes_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
            this.Close();
        }

        private void buttonNo_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        int _printTimeOut = 0;
        private void PrintForm_Load(object sender, EventArgs e)
        {
            IniFile _ini = new IniFile(string.Concat(Application.StartupPath, @"\setting.ini"));
            int time;
            // int 값으로 변환이 가능한 값이면 변환값, 변환이 불가능한 값이면 디폴트 5초)
            _printTimeOut = int.TryParse(_ini.GetIniValue("ENV", "PrintTimeOut"), out time) ? time : 5;
            _Timer.Interval = _printTimeOut * 1000; // n초
            _Timer.Start();
        }

        // 메시지 팝업 열린 후, n초간 액션이 없을 시, 자동 취소처리
        private int _TIMER_CNT = 0;
        private void _Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                Log.AddLog("2차계근 후 출력 자동 취소처리");
                DialogResult = DialogResult.No;
                Close();
            }
            catch (Exception ex)
            {
                Log.AddLog(ex.Message);
            }
        }

        private void PrintForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                _Timer.Stop();
            }
            catch (Exception ex)
            {
                Log.AddLog(ex.Message);
            }
        }

        private void PrintForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //#001
            //Process[] prcList = Process.GetProcessesByName("FxVRAPI_FHS_App");
            //if (prcList.Length != 0)
            //{
            //    prcList[0].Kill();
            //}
        }
    }
}
