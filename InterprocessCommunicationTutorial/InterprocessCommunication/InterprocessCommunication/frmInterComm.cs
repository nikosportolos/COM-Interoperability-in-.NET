using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace InterprocessCommunication
{
    public partial class frmInterComm : Form, IDisposable
    {
        private WindowMessaging WinMsg;

        public frmInterComm()
        {
            InitializeComponent();
            WinMsg = new WindowMessaging();
        }

        private void btSend2VB_Click(object sender, EventArgs e)
        {
            WinMsg.SendMessage();
        }
    }
}
