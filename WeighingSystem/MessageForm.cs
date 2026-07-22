using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WeighingSystem
{
    public partial class MessageForm : Form
    {
        public string FormTitle
        {
            get { return _title; }
            set { _title = value; }
        }
        private string _title = string.Empty;

        public string Message_1
        {
            get { return _message_1; }
            set { _message_1 = value; }
        }
        private string _message_1 = string.Empty;

        public string Message_2
        {
            get { return _message_2; }
            set { _message_2 = value; }
        }
        private string _message_2 = string.Empty;

        public string Message_3
        {
            get { return _message_3; }
            set { _message_3 = value; }
        }
        private string _message_3 = string.Empty;

        public MessageForm()
        {
            InitializeComponent();
        }

        private void MessageForm_Load(object sender, EventArgs e)
        {
            this.Text = FormTitle;

            this.label1.Text = Message_1;
            this.label2.Text = Message_2;
            this.label3.Text = Message_3;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
