using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace checklocal
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemSend += Application_ItemSend;
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {
                string address;


                Outlook.MailItem mail = (Outlook.MailItem)Item;
                Outlook.Recipient bcc;

                if (Properties.Settings.Default.AddBCC != "" && (mail.BCC == null || !mail.BCC.Contains(Properties.Settings.Default.AddBCC)))
                {
                    address = mail.SendUsingAccount.SmtpAddress;
                    string[] BCCaddresses = Properties.Settings.Default.BCCSender.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (BCCaddresses.Length == 0 || BCCaddresses.Contains(address))
                    {
                        bcc = mail.Recipients.Add(Properties.Settings.Default.AddBCC);
                        bcc.Type = (int)Outlook.OlMailRecipientType.olBCC;
                        bcc.Resolve();
                    }
                }

                if (WildcardMatch(mail.SendUsingAccount.SmtpAddress, Properties.Settings.Default.checkSender, false))
                {
                    DialogResult dr = MessageBox.Show(
                        "E-Mail wirklich vom Account (" + Properties.Settings.Default.checkSender + ") verschicken?",
                        "Mail lokal veerschicken?",
                        MessageBoxButtons.OKCancel
                        );
                    if (dr == DialogResult.Cancel) Cancel = true;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace + "\n" + e.Source);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private bool WildcardMatch(String s, String wildcard, bool case_sensitive)
        {
            String pattern = "^" + Regex.Escape(wildcard).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";

            Regex regex;
            if (case_sensitive)
                regex = new Regex(pattern);
            else
                regex = new Regex(pattern, RegexOptions.IgnoreCase);

            return (regex.IsMatch(s));
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
