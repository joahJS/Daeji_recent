using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TextIndicator
{
    public partial class Indicator : UserControl
    {
        #region properities

        delegate void SetConnectCallback();
        delegate void SetStableCallback();
        delegate void SetIndicatorWeightCallback(string value);
        
        public string SystemLanguage
        {
            get { return _language; }
            set { _language = value; }
        }
        private string _language = string.Empty;

        /// <summary>
        /// Indicator Connect 여부 (Lamp Display)
        /// </summary>
        public bool Connect
        {
            get { return _connect; }
            set
            {
                if (_connect != value)
                {
                    _connect = value;

                    SetConnect();
                }
            }
        }
        private bool _connect;

        /// <summary>
        /// Connect Lamp Visible
        /// </summary>
        public bool ConnectVisible
        {
            get { return pictureBoxConnect.Visible; }
            set { pictureBoxConnect.Visible = value; }
        }

        /// <summary>
        /// Indicator 중량 안정화 여부 (Lamp Display)
        /// </summary>
        public bool Stable
        {
            get { return _stable; }
            set
            {
                if (_stable != value)
                {
                    _stable = value;

                    SetStable();
                }
            }
        }
        private bool _stable;

        /// <summary>
        /// Stable Lamp Visible
        /// </summary>
        public bool StableVisible
        {
            get { return pictureBoxStable.Visible; }
            set { pictureBoxStable.Visible = value; }
        }

        /// <summary>
        /// 중량단위 (기본값 = kg)
        /// </summary>
        public string Unit
        {
            get { return _unit; }
            set { _unit = value; }
        }
        private string _unit = "kg";

        /// <summary>
        /// 중량 단위사이 공백여부
        /// </summary>
        public bool UnitSpace
        {
            get { return _space; }
            set { _space = value; }
        }
        private bool _space;

        /// <summary>
        /// Indicator 중량값 (Label Display)
        /// </summary>
        public int Weight
        {
            get { return _weight; }
            set
            {
                if (_weight != value)
                {
                    _weight = value;

                    string display = string.Empty;
                    display = string.Format("{0,6:N0}{1}{2}", _weight, _space ? " " : "", _unit);
                    // display = string.Format("{0}{1}{2}", _weight.ToString("#,##0"), _space ? " " : "", _unit);

                    // 2016.01.19
                    // - 지역설정을 브라질(포어)로 지정하였을 경우 천단위컴마가 소수점으로 표시되는 현상이 발생
                    // - 소수점을 컴마로 강제변경한다.
                    // SetIndicatorWeight(display.Replace(".", ","));

                    // 2016.05.07 
                    // - 언어에 맞게 중량을 명시적으로 표현
                    string strWeight = string.Empty;
                    if (SystemLanguage == "pt-BR")
                        strWeight = display.Replace(',', '.');
                    else
                        strWeight = display.Replace('.', ',');

                    SetIndicatorWeight(strWeight);
                }
            }
        }
        private int _weight = 0;

        /// <summary>
        /// Lamp Margin
        /// </summary>
        public int LampMargin
        {
            get { return _lampMargin; }
            set
            {
                _lampMargin = value;

                OnSizeChanged(null);
            }
        }
        private int _lampMargin;

        /// <summary>
        /// Label Margin
        /// </summary>
        public int LabelMargin
        {
            get { return _labelMargin; }
            set
            {
                _labelMargin = value;

                OnSizeChanged(null);
            }
        }
        private int _labelMargin;

        /// <summary>
        /// Connect, Stable 램프간 간격 (새로)
        /// </summary>
        public int LampGap
        {
            get { return _lampGap; }
            set
            {
                if (value >= 0 && value <= this.Height / 2)
                {
                    _lampGap = value;

                    OnSizeChanged(null);
                }
            }
        }
        private int _lampGap;

        /// <summary>
        /// 라벨 우측 간격
        /// </summary>
        private int WeightRightMargin
        {
            get { return _rightMargin; }
            set { _rightMargin = value; }
        }
        private int _rightMargin;

        /// <summary>
        /// Weight Display Font
        /// </summary>
        public Font WeightFont
        {
            get { return labelWeight.Font; }
            set 
            { 
                labelWeight.Font = value;
                labelWeight.Update();
            }
        }

        private bool _firstLoad = false;

        #endregion

        #region Constructor

        public Indicator()
        {
            InitializeComponent();

            // Default Value ...
            _connect = false;
            _stable = false;
            _unit = "kg";
            _weight = int.MinValue;

            _space = false;

            _lampMargin = 3;
            _labelMargin = 3;
            _lampGap = 14;
            _rightMargin = 20;
        }

        #endregion

        #region Events

        private void Indicator_Load(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                _weight = 0;
                OnSizeChanged(null);
            }
        }

        private void Indicator_SizeChanged(object sender, EventArgs e)
        {
            // 1. connect, stable 위치 및 크기 조정
            // - 새로크기에 맞춤
            int lampSize = (this.Height - (_lampGap + (_lampMargin * 2))) / 2;

            // (1) Connect
            pictureBoxConnect.Location = new Point(_lampMargin, _lampMargin);
            pictureBoxConnect.Size = new Size(lampSize, lampSize);
            // (2) Stable
            pictureBoxStable.Location = new Point(_lampMargin, lampSize + _lampMargin + _lampGap);
            pictureBoxStable.Size = new Size(lampSize, lampSize);

            // 2. 중량 Label
            labelWeight.Location = new Point(lampSize + (_lampMargin * 3), _labelMargin);
            int wSize = this.Width - labelWeight.Location.X - (_rightMargin + _labelMargin);
            int hSize = this.Height - (_labelMargin * 2);
            labelWeight.Size = new Size(wSize, hSize);

        }

        #endregion

        #region Methods

        #endregion

        #region Implementations

        private void SetConnect()
        {
            if (pictureBoxConnect.InvokeRequired)
            {
                SetConnectCallback d = new SetConnectCallback(SetConnect);
                this.Invoke(d);
            }
            else
            {
                if (!_connect)
                    pictureBoxConnect.BackgroundImage = Properties.Resources.connect_false;
                else
                    pictureBoxConnect.BackgroundImage = Properties.Resources.connect_true;
            }
        }

        private void SetStable()
        {
            if (pictureBoxStable.InvokeRequired)
            {
                SetStableCallback d = new SetStableCallback(SetStable);
                this.Invoke(d);
            }
            else
            {
                if (!_stable)
                    pictureBoxStable.BackgroundImage = Properties.Resources.steady_false;
                else
                    pictureBoxStable.BackgroundImage = Properties.Resources.steady_true;
            }
        }

        private void SetIndicatorWeight(string value)
        {
            if (labelWeight.InvokeRequired)
            {
                SetIndicatorWeightCallback d = new SetIndicatorWeightCallback(SetIndicatorWeight);
                this.Invoke(d, new object[] { value });
            }
            else
            {
                labelWeight.Text = value;
            }
        }

        #endregion
    }
}
