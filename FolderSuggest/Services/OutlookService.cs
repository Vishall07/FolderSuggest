using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using FolderSuggest.Models;

namespace FolderSuggest.Services
{
    public class OutlookService
    {
        private readonly Outlook.Application _outlookApp;

        public OutlookService(Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
        }

        public List<EmailItem> GetEmailsFromFolders()
        {
            var emails = new List<EmailItem>();
            try
            {
                var ns = _outlookApp.GetNamespace("MAPI");
                var rootFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                CollectEmailsFromFolder(rootFolder, emails);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching emails: {ex.Message}");
            }
            return emails;
        }

        private void CollectEmailsFromFolder(Outlook.MAPIFolder folder, List<EmailItem> emails)
        {
            try
            {
                var items = folder.Items;
                items.Sort("[ReceivedTime]", true);

                foreach (object item in items)
                {
                    if (item is Outlook.MailItem mailItem)
                    {
                        var email = new EmailItem
                        {
                            Subject = mailItem.Subject,
                            From = mailItem.SenderName,
                            ReceivedTime = mailItem.ReceivedTime,
                            FolderName = folder.Name,
                            Body = mailItem.Body
                        };
                        emails.Add(email);
                    }
                }

                foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                {
                    CollectEmailsFromFolder(subFolder, emails);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing folder: {ex.Message}");
            }
        }
    }
}
