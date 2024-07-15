using Microsoft.Office.Interop.Outlook;
using System.Configuration;

namespace VSTOOutlookAddInReadEmbeddedTables
{
    public partial class ThisAddIn
    {
        private string FromEmail;
        private string ToEmail;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            FromEmail = ConfigurationManager.AppSettings["FromEmail"];
            ToEmail = ConfigurationManager.AppSettings["ToEmail"];

            Application.NewMailEx += Application_NewMailEx;

            ProcessUnreadEmails();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void Application_NewMailEx(string EntryIDCollection)
        {
            try
            {
                MailItem newMail = (MailItem)Application.Session.GetItemFromID(EntryIDCollection);
                if (newMail != null && newMail.SenderName == FromEmail)
                {
                    string[] values = HtmlTableParser.ExtractValuesFromHtml(newMail.HTMLBody);

                    // Add any additional processing here
                    newMail.UnRead = false;
                    newMail.Save();
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error handling new mail: " + ex.Message);
            }
        }
        private void ProcessUnreadEmails()
        {
            MAPIFolder inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            Items unreadItems = inbox.Items.Restrict("[Unread]=true");

            foreach (MailItem mail in unreadItems)
            {
                if (mail != null && mail.SenderName == FromEmail && mail.To == ToEmail)
                {
                    try
                    {
                        string[] values = HtmlTableParser.ExtractValuesFromHtml(mail.HTMLBody);

                        // Optionally, mark as read (if desired)
                        mail.UnRead = false;
                        mail.Save();

                        // Perform any additional processing here
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Error processing mail: " + ex.Message);
                    }
                }
            }
        }

        


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
