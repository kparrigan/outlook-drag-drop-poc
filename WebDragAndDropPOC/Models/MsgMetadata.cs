using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebDragAndDropPOC.Models
{
    public sealed class MsgMetadata
    {
        public MsgMetadata()
        {
            Attachments = new List<string>();
        }

        public string FileName { get; set; }
        public string Sender { get; set; }
        public List<string> RecipientsTo { get; set; }
        public List<string> RecipientsCc { get; set; }
        public List<string> RecipientsBcc { get; set; }
        public DateTime? SentDateTime { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<string> Attachments { get; set; }
    }
}