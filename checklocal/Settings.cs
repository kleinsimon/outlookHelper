using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace checklocal
{
    public partial class Settings : Form
    {
        Ribbon1 parent = null;
        public Settings(Ribbon1 prn)
        {
            InitializeComponent();
            parent = prn;

            textCheck.Text = parent.checkSender;
            textBCC.Text = parent.AddBCC;
            textBCCSender.Text = parent.BCCSender;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            parent.checkSender = textCheck.Text;
            parent.AddBCC = textBCC.Text;
            parent.BCCSender = textBCCSender.Text;

            this.Close();
        }
    }
}
