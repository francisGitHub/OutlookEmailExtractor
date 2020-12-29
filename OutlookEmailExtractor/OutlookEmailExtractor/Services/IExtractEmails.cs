using System.Collections.Generic;
using OutlookEmailExtractor.Model;

namespace OutlookEmailExtractor.Services
{
    public interface IExtractEmails
    {
        IList<Email> GetEmails();
    }
}
