using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using WebDragAndDropPOC.Models;
using msgReader = MsgReader.Outlook;

namespace WebDragAndDropPOC.Middleware
{
    public sealed class MsgParser : IMsgParser
    {
        private readonly HashSet<string> _attachmentFilterSet;

        public MsgParser()
        {
            _attachmentFilterSet = new HashSet<string>();
        }

        public MsgParser(HashSet<string> attachmentFilters)
        {   
            if (attachmentFilters == null)
            {
                throw new ArgumentNullException(nameof(attachmentFilters));                
            }

            _attachmentFilterSet = new HashSet<string>(attachmentFilters);
        }

        public MsgMetadata Parse(string localFileName, string originalFileName)
        {
            if (string.IsNullOrWhiteSpace(localFileName))
            {
                throw new ArgumentException("Local file name must have a value.", nameof(localFileName));
            }

            if (!File.Exists(localFileName))
            {
                throw new FileNotFoundException($"{localFileName} does not exist.", nameof(localFileName));
            }

            if (string.IsNullOrWhiteSpace(originalFileName))
            {
                throw new ArgumentException("Original file name must have a value.", nameof(originalFileName));
            }

            var metadata = new MsgMetadata();

            using (var msg = new msgReader.Storage.Message(localFileName))
            {
                metadata.FileName = originalFileName;
                metadata.Sender = msg.Sender.Email;
                metadata.SentDateTime = msg.SentOn;

                var recipientsRaw = msg.GetEmailRecipients(msgReader.Storage.Recipient.RecipientType.To, false, false);
                metadata.RecipientsTo = GetParticipantList(recipientsRaw);

                var recipientsCCRaw = msg.GetEmailRecipients(msgReader.Storage.Recipient.RecipientType.Cc, false, false);
                metadata.RecipientsCc = GetParticipantList(recipientsCCRaw);

                var recipientsBCCRaw = msg.GetEmailRecipients(msgReader.Storage.Recipient.RecipientType.Bcc, false, false);
                metadata.RecipientsBcc = GetParticipantList(recipientsBCCRaw);

                metadata.Subject = msg.Subject;
                metadata.Body = msg.BodyText;

                var attachments = msg.Attachments;
                if (attachments.Count != 0)
                {
                    for (int i = 0, count = attachments.Count; i < count; i++)
                    {
                        if (attachments[i] is msgReader.Storage.Attachment attachment)
                        {
                            // If we have attachment types we don't care about, filter those attachments out. 
                            if (_attachmentFilterSet.Any() && _attachmentFilterSet.Contains(Path.GetExtension(attachment.FileName)))
                            {
                                continue;
                            }

                            metadata.Attachments.Add(attachment.FileName);
                        }
                    }
                }
            }

            return metadata;
        }

        private List<string> GetParticipantList(string participants)
        {
            if (string.IsNullOrWhiteSpace(participants))
            {
                return new List<string>();
            }

            return participants.Split(';').ToList();
        }
     }
}