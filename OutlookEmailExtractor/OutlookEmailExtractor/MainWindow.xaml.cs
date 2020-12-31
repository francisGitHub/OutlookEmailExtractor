using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using CsvHelper;
using Microsoft.Office.Interop.Outlook;
using OutlookEmailExtractor.Model;
using OutlookEmailExtractor.Services;
using Exception = System.Exception;

namespace OutlookEmailExtractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly IExtractEmails _emailExtractionService;
        private string SaveDirectory = "C:\\Code\\Emails";
        public MainWindow(
            IExtractEmails emailExtractionService)
        {
            _emailExtractionService = emailExtractionService;
            InitializeComponent();
        }

        private void ExtractEmails(object sender, RoutedEventArgs e)
        {
            var emails = _emailExtractionService.GetEmails();
            var orderedEmails = emails.OrderBy(x => x.ReceivedDateTime);
            int count = 0;
            var indexFormat = new IndexFormat(4, 2);

            var csvRows = new List<CsvRow>();

            try
            {
                Directory.CreateDirectory(SaveDirectory);

                foreach (var item in orderedEmails)
                {
                    int attachmentIdentifier = 0;
                    count++;
                    
                    var emailIndex = indexFormat.GetIndexNumber(count, 0);

                    var emailSavePath = Path.Combine(SaveDirectory, $"{emailIndex}.msg");
                    item.MailItem.SaveAs($"{emailSavePath}", OlSaveAsType.olMSG);

                    var csvEmailRowItem = new CsvRow
                    {
                        Index = emailIndex,
                        Description = item.Subject,
                        Date = item.ReceivedDateTime.ToShortDateString(),
                        Time = item.ReceivedDateTime.ToString("HH:mm").PadLeft(4,'0')
                    };

                    csvRows.Add(csvEmailRowItem);

                    if (item.Attachments != null && item.Attachments.Any())
                    {
                        foreach (var attachment in item.Attachments)
                        {


                            attachmentIdentifier++;
                            var attachmentIndex = indexFormat.GetIndexNumber(count, attachmentIdentifier);

                            var attachmentSavePath =
                                Path.Combine(SaveDirectory, $"{attachmentIndex} {attachment.FileName}");
                            attachment.SaveAsFile(attachmentSavePath);

                            var csvAttachmentRowItem = new CsvRow
                            {
                                Index = attachmentIndex, Description = attachment.FileName
                            };

                            csvRows.Add(csvAttachmentRowItem);
                        }
                    }
                }

                var csvFileName = "ExtractedEmails.csv";
                var csvFilePath = Path.Combine(SaveDirectory, csvFileName);

                using (var textWriter = new StreamWriter(csvFilePath))
                {
                    var writer = new CsvWriter(textWriter, CultureInfo.InvariantCulture);
                    writer.Configuration.Delimiter = ",";
                    writer.WriteRecords(csvRows);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw;
            }
        }
    }
}
