using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using OutlookEmailExtractor.Model;
using Action = System.Action;

namespace OutlookEmailExtractor.Services.Impl
{
    public class EmailExtractionService : IExtractEmails
    {
        public IList<Email> GetEmails()
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items items = null;

            List<Email> emails = new List<Email>();

            try
            {
                outlookApplication = new Application();
                inboxFolder = outlookApplication.Session.PickFolder();
                items = inboxFolder.Items;

                foreach (var item in items)
                {
                    if (item is MailItem mailItem)
                    {
                        var email = new Email
                        {
                            MailItem = mailItem,
                            SenderEmailAddress = GetSenderEmailAddress(mailItem),
                            To = mailItem.To,
                            Subject = mailItem.Subject,
                            ReceivedDateTime = mailItem.ReceivedTime,
                        };

                        var attachments = mailItem.Attachments;

                        if (attachments.Count != 0)
                        {
                            email.Attachments = new List<Attachment>();
                            foreach (Attachment attachment in mailItem.Attachments)
                            {
                                var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");

                                //To ignore embedded attachments -
                                if (flags != 4)
                                {
                                    // As per present understanding - If rtF mail attachment comes here - and the embeded image is treated as attachment then Type value is 6 and ignore it
                                    if ((int)attachment.Type != 6)
                                    {
                                        string[] extensionsArray = { ".pdf", ".doc", ".xls", ".ppt", ".vsd", ".zip", ".rar", ".txt", ".csv", ".proj", ".docx", ".obr" };

                                        if (extensionsArray.Any(attachment.FileName.Contains))
                                        {
                                            email.Attachments.Add(attachment);
                                        }
                                    }
                                }
                            }
                        }

                        emails.Add(email);
                    }
                    else
                    {
                        Debug.WriteLine($"Item is not a {nameof(MailItem)}");
                    }
                }
            }
            //Error handler.
            catch (System.Exception e)
            {
                Console.WriteLine("{0} Exception caught: ", e);
            }
            finally
            {
                ReleaseComObject(items);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }

            return emails;
        }

        private void ReleaseComObject(object obj)
        {
            if (obj == null)
            {
                return;
            }

            Marshal.ReleaseComObject(obj);
            obj = null;
        }

        private string GetSenderEmailAddress(MailItem mail)
        {
            AddressEntry sender = mail.Sender;
            string senderEmailAddress = "";

            if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null)
                {
                    senderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else
            {
                senderEmailAddress = mail.SenderEmailAddress;
            }

            return senderEmailAddress;
        }
    }
}
