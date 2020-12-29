using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace OutlookEmailExtractor.Model
{
    public class Email
    {
        public MailItem MailItem { get; set; }
        public string Subject { get; set; }
        public DateTime ReceivedDateTime { get; set; }
        public string SenderEmailAddress { get; set; }
        public string To { get; set; }
        public List<Attachment> Attachments { get; set; }
    }
}
