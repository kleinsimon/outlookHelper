using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace checklocal
{
    public partial class Ribbon1
    {
        public string BCCSender
        {
            get
            {
                return Properties.Settings.Default.BCCSender;
            }
            set
            {
                Properties.Settings.Default.BCCSender = value;
                Properties.Settings.Default.Save();
            }
        }
        public string AddBCC
        {
            get
            {
                return Properties.Settings.Default.AddBCC;
            }
            set
            {
                Properties.Settings.Default.AddBCC = value;
                Properties.Settings.Default.Save();
                checkIfSet();
            }
        }
        public string checkSender
        {
            get
            {
                return Properties.Settings.Default.checkSender;
            }
            set
            {
                Properties.Settings.Default.BCCSender = value;
                Properties.Settings.Default.Save();
            }
        }
        public bool doAddBCC
        {
            get
            {
                return Properties.Settings.Default.doAddBCC;
            }
            set
            {
                Properties.Settings.Default.doAddBCC = value;
                Properties.Settings.Default.Save();
                checkBoxDoBcc.Checked = value;
            }
        }
        public bool doBCCFeedback
        {
            get
            {
                return Properties.Settings.Default.doBCCFeedback;
            }
            set
            {
                Properties.Settings.Default.doBCCFeedback = value;
                Properties.Settings.Default.Save();
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            checkIfSet();
            checkBoxDoBcc.Checked = doAddBCC;
        }

        public void checkIfSet()
        {
            if (AddBCC == "" || !AddBCC.Contains('@'))
            {
                checkBoxDoBcc.Enabled = false;
                checkBoxDoBcc.SuperTip = "Bitte legen Sie in den Einstellungen eine BCC-Adresse fest";
            }
            else
            {
                checkBoxDoBcc.Checked = false;
                checkBoxDoBcc.Enabled = true;
                checkBoxDoBcc.SuperTip = "";
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Settings setting = new Settings(this);
            setting.ShowDialog();
        }

        private void checkBoxDoBcc_Click(object sender, RibbonControlEventArgs e)
        {
            doAddBCC = !doAddBCC;
        }
    }
}
