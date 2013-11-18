using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

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
                string address, tld;
                Outlook.MailItem mail = (Outlook.MailItem)Item;

                address = mail.SendUsingAccount.SmtpAddress;
                tld = address.Substring(address.LastIndexOf("."));
                if (tld == ".local")
                {
                    DialogResult dr = MessageBox.Show("E-Mail wirklich vom lokalen Account (*@wtech.local) verschicken?", "Mail lokal veerschicken?", MessageBoxButtons.OKCancel);
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
