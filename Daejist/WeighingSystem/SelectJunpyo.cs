using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static WeighingSystem.MainForm;

namespace WeighingSystem
{
    public partial class SelectJunpyo : Form
    {
        public SelectJunpyo()
        {
            InitializeComponent();
            _PARAM = new Dictionary<int, string>();
        }

        public MainForm _FORM;

        public DataTable _DT_RST;
        Dictionary<int, string> _PARAM;
        public WeighingType _Gubun;

        private void SelectJunpyo_Load(object sender, EventArgs e)
        {
            int iCnt = 0;
            //Dictionary 세팅 -> 리스트에서 선택한 인덱스값으로 JunpyoID를 가져오기위함
            foreach (DataRow row in _DT_RST.Rows)
            {
                string sKERATYPE = row["KERATYPE"]?.ToString();
                if(_Gubun == WeighingType.In)
                {
                    if (sKERATYPE.Equals("입고"))
                    {
                        _PARAM.Add(iCnt++, row["JUNPYOID"]?.ToString());
                    }
                }
                else if(_Gubun == WeighingType.Out)
                {
                    if (sKERATYPE.Equals("출고"))
                    {
                        _PARAM.Add(iCnt++, row["JUNPYOID"]?.ToString());
                    }
                }
            }

            foreach (DataRow row in _DT_RST.Rows)
            {
                string sKERATYPE = row["KERATYPE"]?.ToString();
                string sJ_DATE = row["J_DATE"]?.ToString();
                string sJ_BNUM = row["J_BNUM"]?.ToString();
                string sFIRST_TIME = row["FIRSTTIME"]?.ToString();
                string sSECOND_TIME = row["SECONDTIME"]?.ToString();
                string sFIRST_WEIGHT = row["FIRSTWEIGHT"]?.ToString();
                string sSECOND_WEIGHT = row["SECONDWEIGHT"]?.ToString();
                string sGUBUN1 = row["GUBUN1"]?.ToString();
                
                if (_Gubun == WeighingType.In)
                {
                    if (sKERATYPE.Equals("입고"))
                    {
                        string dt = Convert.ToDateTime(sSECOND_TIME).ToString("HH시mm분");
                        string sItem = string.Format("차번 : {0} / 계근시간 : {1} / 대지중량 : {2}"
                                                    , sJ_BNUM
                                                    , dt
                                                    , sSECOND_WEIGHT);

                        listBoxCustomers.Items.Add(sItem);
                    }
                }
                else if (_Gubun == WeighingType.Out)
                {
                    if (sKERATYPE.Equals("출고"))
                    {
                        string dt = Convert.ToDateTime(sFIRST_TIME).ToString("HH시mm분");
                        string sItem = string.Format("차번 : {0} / 계근시간 : {1} / 공차중량 : {2}"
                                                    , sJ_BNUM
                                                    , dt
                                                    , sSECOND_WEIGHT);

                        listBoxCustomers.Items.Add(sItem);
                    }
                }
            }
        }

        private void listBoxCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(listBoxCustomers.SelectedIndex >= 0)
            {
                BtnSelect.BackColor = Color.Blue;
                BtnSelect.ForeColor = Color.White;
                BtnSelect.Enabled = true;
            }
            else
            {
                BtnSelect.BackColor = Color.Red;
                BtnSelect.ForeColor = Color.Yellow;
                BtnSelect.Enabled = false;
            }
        }

        private void listBoxCustomers_DoubleClick(object sender, EventArgs e)
        {
            if (e.Equals(listBoxCustomers))
            {
                SelectItem();
            }
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            if (listBoxCustomers.SelectedIndex < 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "리스트의 계근실적을 선택하세요.");
                return;
            }

            try
            {
                SelectItem();
            }
            catch(Exception ex)
            {
                MessageBox.Show(new Form { TopMost = true }, ex.Message);
            }
        }

        private void SelectItem()
        {
            _FORM._selected_JunpyoID = _PARAM[listBoxCustomers.SelectedIndex];
            DialogResult = DialogResult.OK;
            Dispose();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Dispose();
        }
    }
}
