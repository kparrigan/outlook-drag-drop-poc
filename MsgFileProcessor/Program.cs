using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using interop = Microsoft.Office.Interop.Outlook;
using msgReader = MsgReader.Outlook;

namespace MsgFileProcessor
{
    class Program
    {
        private static interop.Application _outlookApplication;

        static void Main(string[] args)
        {
            const string outputDirectory = @"C:\Git\poc\OutlookDragAndDropPOC\data\";
            const string msgFileName = @"C:\Git\poc\OutlookDragAndDropPOC\data\test.msg";
            //ProcessWithInteropApi(msgFileName, outputDirectory);
            ProcessWithMsgReader(msgFileName, outputDirectory);


            Console.ReadKey();
        }

        private static void ProcessWithMsgReader(string msgFileName, string outputDirectory)
        {
            using (var msg = new msgReader.Storage.Message(msgFileName))
            {
                var from = msg.Sender;
                var sentOn = msg.SentOn;
                var recipientsTo = msg.GetEmailRecipients(msgReader.Storage.Recipient.RecipientType.To, false, false);
                var recipientsCc = msg.GetEmailRecipients(msgReader.Storage.Recipient.RecipientType.Cc, false, false);
                var subject = msg.Subject;
                var htmlBody = msg.BodyHtml;

                var attachments = msg.Attachments;
                if (attachments.Count != 0)
                {
                    for (int i = 0, count = attachments.Count; i < count; i++)
                    {
                        var attachment = attachments[i] as msgReader.Storage.Attachment;
                        if (attachment != null)
                        {
                            using (var streamWriter = new StreamWriter(Path.Combine(outputDirectory, attachment.FileName)))
                            {
                                streamWriter.BaseStream.Write(attachment.Data, 0, attachment.Data.Length);
                            }
                        }
                    }
                }
            }
        }

        private static void ProcessWithInteropApi(string msgFileName, string outputDirectory)
        {
            GetCurrentOutlookApplicationInstance();
            var mailItem = _outlookApplication.Session.OpenSharedItem(msgFileName);
            Console.WriteLine(mailItem.SenderEmailAddress);

            var attachments = mailItem.Attachments;

            if (attachments.Count != 0)
            {
                for (int i = 1, count = attachments.Count; i <= count; i++)
                {
                    var attachment = attachments[i];
                    attachment.SaveAsFile(Path.Combine(outputDirectory, attachment.FileName));
                }
            }
        }

        private static void GetCurrentOutlookApplicationInstance()
        {
            try
            {
                if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                {
                    _outlookApplication = Marshal.GetActiveObject("Outlook.Application") as interop.Application;

                    if (_outlookApplication == null)
                    {
                        //TODO show failure to initialize
                    }
                }
            }
            catch (Exception ex)
            {
                //TODO show exception
                throw new Exception("temp");
            }
        }
    }
}
