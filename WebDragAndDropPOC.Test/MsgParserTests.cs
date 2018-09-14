using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WebDragAndDropPOC.Middleware;

namespace WebDragAndDropPOC.Test
{
    [TestClass]
    public class MsgParserTests
    {
        [TestMethod]
        public void CanParseMsg()
        {
            const string fileName = "test.msg";
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var filePath = Path.Combine(directory, "Data", fileName);
            var parser = new MsgParser();

            var metadata = parser.Parse(filePath, fileName);

            Assert.AreEqual("This is a test email", metadata.Subject);
            Assert.AreEqual("This is just a test\r\n", metadata.Body);
            Assert.AreEqual("kyle.parrigan@gmail.com", metadata.Sender);
            Assert.AreEqual(2, metadata.RecipientsTo.Count);
            Assert.AreEqual(1, metadata.RecipientsCc.Count);
            Assert.AreEqual(0, metadata.RecipientsBcc.Count);
            Assert.AreEqual(2, metadata.Attachments.Count);
        }

        [TestMethod]
        public void CanFilterAttachments()
        {
            const string fileName = "filter_test.msg";
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var filePath = Path.Combine(directory, "Data", fileName);
            var filters = new HashSet<string>() { ".gif" };
            var parser = new MsgParser(filters);

            var metadata = parser.Parse(filePath, fileName);

            Assert.AreEqual(1, metadata.Attachments.Count);
        }
    }
}
