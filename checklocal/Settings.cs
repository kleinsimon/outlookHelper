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
        public Settings()
        {
            InitializeComponent();
            textCheck.Text = Properties.Settings.Default.checkSender;
            textBCC.Text = Properties.Settings.Default.AddBCC;
            textBCCSender.Text = Properties.Settings.Default.BCCSender;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.checkSender = textCheck.Text;
            Properties.Settings.Default.AddBCC = textBCC.Text;
            Properties.Settings.Default.BCCSender = textBCCSender.Text;

            Properties.Settings.Default.Save();

            this.Close();
        }
    }
}
