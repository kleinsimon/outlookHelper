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
            this.Application.Inspectors.NewInspector += Inspectors_NewInspector;
        }

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            if (!Properties.Settings.Default.doAddBCC) return;
            try
            {
                Outlook.MailItem mail = (Outlook.MailItem)Inspector.CurrentItem;
                if (mail.Sent)
                    return;
                mail.PropertyChange += mail_PropertyChange;
                UpdateBccInMail(mail);
            }
            catch { }
        }

        private void UpdateBccInMail(Outlook.MailItem mail)
        {
            if (!Properties.Settings.Default.doAddBCC) return;
            Outlook.Recipient bcc;
            string[] BCCaddresses = Properties.Settings.Default.BCCSender.Trim().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            string address = mail.SendUsingAccount.SmtpAddress;

            if (BCCaddresses.Length == 0 || BCCaddresses.Contains(address))
            {
                if (RecipientsContainAddress(mail.Recipients, Properties.Settings.Default.AddBCC.Trim()))
                    return;
                bcc = mail.Recipients.Add(Properties.Settings.Default.AddBCC.Trim());
                bcc.Type = (int)Outlook.OlMailRecipientType.olBCC;
                bcc.Resolve();
            }
            else
            {
                removeRecipient(mail.Recipients, Properties.Settings.Default.AddBCC);
            }
        }

        void removeRecipient(Outlook.Recipients recipients, string address)
        {
            bool res = true;
            while (res)
            {
                res = removeRecipientCall(recipients, address);
            }
        }

        bool removeRecipientCall(Outlook.Recipients recipients, string address)
        {
            int rm = -1;

            foreach (Outlook.Recipient rc in recipients)
            {
                if (rc.Address.Trim() == address.Trim())
                {
                    rm = rc.Index;
                }
            }

            if (rm == -1)
            {
                return false;
            }
            else
            {
                recipients.Remove(rm);
                return true;
            }
        }

        bool RecipientsContainAddress(Outlook.Recipients recipients, string address)
        {
            foreach (Outlook.Recipient rc in recipients)
            {
                if (rc.Address.Trim() == address.Trim())
                    return true;
            }

            return false;
        }

        void mail_PropertyChange(string Name)
        {
            if (Name == "SendUsingAccount")
            {
                UpdateCurrentMail();
            }
        }

        public void UpdateCurrentMail()
        {
            if (!Properties.Settings.Default.doAddBCC) return;
            try
            {
                Outlook.MailItem mail = this.Application.ActiveInspector().CurrentItem;
                string address = mail.SendUsingAccount.SmtpAddress;
                UpdateBccInMail(mail);
            }
            catch { }
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {
                string address;

                Outlook.MailItem mail = (Outlook.MailItem)Item;
                address = mail.SendUsingAccount.SmtpAddress;

                //if (Properties.Settings.Default.doAddBCC && Properties.Settings.Default.AddBCC != "" && (mail.BCC == null || !mail.BCC.Contains(Properties.Settings.Default.AddBCC)))
                //{          
                //    UpdateBccInMail(mail);
                //    if (Properties.Settings.Default.doAddBCC)
                //        MessageBox.Show("Die Addresse " + Properties.Settings.Default.AddBCC + " wurde als BCC hinzugefügt");
                //}

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
            catch
            {
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
