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
using System.IO.Ports;

namespace AccAdm
{
    public partial class PopUpSerialTcpIp : DevExpress.XtraEditors.XtraForm
    {


        public PopUpSerialTcpIp()
        {
            InitializeComponent();

            CboComPort.Properties.BeginUpdate();
            foreach (string comport in SerialPort.GetPortNames())
            CboComPort.Properties.Items.Add(comport);
            CboComPort.Properties.EndUpdate();
            CboComPort.Select();
            CboComPort.SelectedIndex = 0;
        }

        private void CboComPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            serialPort1.PortName = CboComPort.SelectedItem.ToString();
        }

        private void CboBaudRate_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CboBaudRate.SelectedIndex)
            {
                case 0:
                    serialPort1.BaudRate = 9600; break;
                case 1:
                    serialPort1.BaudRate = 14400; break;
                case 2:
                    serialPort1.BaudRate = 19200; break;
                case 3:
                    serialPort1.BaudRate = 38400; break;
                case 4:
                    serialPort1.BaudRate = 57600; break;
                case 5:
                    serialPort1.BaudRate = 115200; break;
                case 6:
                    serialPort1.BaudRate = 128000; break;
                default:
                    serialPort1.BaudRate = 115200; break;
            }
        }

        private void CboDataBits_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CboDataBits.SelectedIndex)
            {
                case 0:
                    serialPort1.DataBits = 8; break;
                case 1:
                    serialPort1.DataBits = 7; break;
                default:
                    serialPort1.DataBits = 8; break;
            }
        }

        private void CboParity_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CboParity.SelectedIndex)
            {
                case 0:
                    serialPort1.Parity = Parity.Even; break;
                case 1:
                    serialPort1.Parity = Parity.Mark; break;
                case 2:
                    serialPort1.Parity = Parity.None; break;
                case 3:
                    serialPort1.Parity = Parity.Odd; break;
                case 4:
                    serialPort1.Parity = Parity.Space; break;
                default:
                    serialPort1.Parity = Parity.None; break;
            }
        }

        private void CboStopBits_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CboStopBits.SelectedIndex)
            {
                case 0:
                    CboStopBits.Text = "이 값은 지원되지 않습니다."; break;
                case 1:
                    serialPort1.StopBits = StopBits.One; break;
                case 2:
                    serialPort1.StopBits = StopBits.OnePointFive; break;
                case 3:
                    serialPort1.StopBits = StopBits.Two; break;
                default:
                    serialPort1.StopBits = StopBits.One; break;
            }
        }

        public string cComportNum;
        public Parity cParity;
        public StopBits cStopBits;
        public int ciBaudRate;
        public int ciDataBit;
        public string ciTcpip;
        public string ciPortNum;

        private void BtnPortOpen_Click(object sender, EventArgs e)
        {
            //tcp/ip
            string sTcpIp = TxtTcpIP.EditValue?.ToString();
            string sPortNum = TxtPortNum.EditValue?.ToString();

            //시리얼포트
            string sBaudRate = CboBaudRate.EditValue.ToString();
            string sDataBit = CboDataBits.EditValue.ToString();

            cComportNum = CboComPort.SelectedItem.ToString();
            ciBaudRate = Convert.ToInt32(sBaudRate);
            ciDataBit = Convert.ToInt32(sDataBit);

            ciTcpip = sTcpIp;
            ciPortNum = sPortNum;


            switch (CboStopBits.SelectedIndex)
            {
                case 0:
                    CboStopBits.Text = "이 값은 지원되지 않습니다."; break;
                case 1:
                    cStopBits = StopBits.One; break;
                case 2:
                    cStopBits = StopBits.OnePointFive; break;
                case 3:
                    cStopBits = StopBits.Two; break;
                default:
                    cStopBits = StopBits.One; break;
            }

            switch (CboParity.SelectedIndex)
            {
                case 0:
                    cParity = Parity.Even; break;
                case 1:
                    cParity = Parity.Mark; break;
                case 2:
                    cParity = Parity.None; break;
                case 3:
                    cParity = Parity.Odd; break;
                case 4:
                    cParity = Parity.Space; break;
                default:
                    cParity = Parity.None; break;
            }
            
            this.DialogResult = DialogResult.OK;

        }

        private void EnableControls(bool value)
        {
            if (value == true)
            {
                CboComPort.Enabled = true;
                CboBaudRate.Enabled = true;
                CboDataBits.Enabled = true;
                CboParity.Enabled = true;
                CboStopBits.Enabled = true;
            }
            else
            {
                CboComPort.Enabled = false;
                CboBaudRate.Enabled = false;
                CboDataBits.Enabled = false;
                CboParity.Enabled = false;
                CboStopBits.Enabled = false;
            }
        }

        private void PopUpSerialTcpIp_Load(object sender, EventArgs e)
        {
            //CboComPort.Properties.Items.Add("COM1");
            //CboComPort.Properties.Items.Add("COM2");
            //CboComPort.Properties.Items.Add("COM3");
            //CboComPort.Properties.Items.Add("COM4");
            //CboComPort.SelectedIndex = 2;
        }
    }
}