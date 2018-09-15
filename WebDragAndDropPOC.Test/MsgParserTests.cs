using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WebDragAndDropPOC.Middleware;
using CustomAttributes;

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

        [TestMethod]
        [CustomExpectedException(typeof(FileNotFoundException), ExpectedExceptionMessage = "foo does not exist.")]
        public void CanHandleInvalidLocalFileName()
        {
            var parser = new MsgParser();

            var metadata = parser.Parse("foo", "bar.txt");
        }

        [TestMethod]
        [CustomExpectedException(typeof(ArgumentException), ParameterName = "localFileName")]
        public void CanHandleEmptyLocalFileName()
        {
            var parser = new MsgParser();

            var metadata = parser.Parse(string.Empty, "bar.txt");
        }

        [TestMethod]
        [CustomExpectedException(typeof(ArgumentException), ParameterName = "originalFileName")]
        public void CanHandleEmptyOriginalFileName()
        {
            const string fileName = "filter_test.msg";
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var filePath = Path.Combine(directory, "Data", fileName);
            var parser = new MsgParser();

            var metadata = parser.Parse(filePath, string.Empty);
        }
    }
}
