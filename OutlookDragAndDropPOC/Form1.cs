using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Zhwang.SuperNotifyIcon;
using interop = Microsoft.Office.Interop.Outlook;

namespace OutlookDragAndDropPOC
{
    public partial class Form1 : Form
    {
        private interop.Application _outlookApplication;
        private SuperNotifyIcon _notifyIcon;

        public Form1()
        {
            InitializeComponent();
            InitializeNotifyIcon();
            GetCurrentOutlookApplicationInstance();

            AllowDrop = true;
            DragEnter += new DragEventHandler(Form1_DragEnter);
            DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {

            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                //TODO - do we want to support multiple message drops?
                if (_outlookApplication.ActiveExplorer().Selection.Count > 0 &&
                    _outlookApplication.ActiveExplorer().Selection[1] is interop.MailItem mailItem) // Interop collections have 1-based indices. Because fuck you, that's why.
                {
                    txtSender.Text = mailItem.SenderEmailAddress;

                    var recipients = mailItem.Recipients;
                    for(int i = 1, count = recipients.Count; i <= count; i++)
                    {
                        var address = GetRecipientSMTPAddress(recipients[i]);
                        var temp = recipients[i].Address;

                        if (txtRecipients.Text.Length == 0)
                        {
                            txtRecipients.Text = address;
                        }
                        else
                        {
                            txtRecipients.AppendText(Environment.NewLine + address);
                        }
                    }

                    txtSubject.Text = mailItem.Subject;
                    txtBody.Text = mailItem.Body;

                    var attachments = mailItem.Attachments;

                    if (attachments.Count != 0)
                    {
                        for (int i = 1, count = attachments.Count; i <= count; i++)
                        {
                            var fileName = attachments[i].FileName;                            

                            if (txtAttachments.Text.Length == 0)
                            {
                                txtAttachments.Text = fileName;
                            }
                            else
                            {
                                txtAttachments.AppendText(Environment.NewLine + fileName);
                            }
                        }
                    }
                }
            }
        }

        private void Icon_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void Icon_DragDrop(object sender, DragEventArgs e)
        {

            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                //TODO - do we want to support multiple message drops?
                if (_outlookApplication.ActiveExplorer().Selection.Count > 0 &&
                    _outlookApplication.ActiveExplorer().Selection[1] is interop.MailItem mailItem) // Interop collections have 1-based indices. Because fuck you, that's why.
                {
                    txtSender.Text = mailItem.SenderEmailAddress;

                    var recipients = mailItem.Recipients;
                    for (int i = 1, count = recipients.Count; i <= count; i++)
                    {
                        var address = GetRecipientSMTPAddress(recipients[i]);
                        var temp = recipients[i].Address;

                        if (txtRecipients.Text.Length == 0)
                        {
                            txtRecipients.Text = address;
                        }
                        else
                        {
                            txtRecipients.AppendText(Environment.NewLine + address);
                        }
                    }

                    txtSubject.Text = mailItem.Subject;
                    txtBody.Text = mailItem.Body;

                    var attachments = mailItem.Attachments;

                    if (attachments.Count != 0)
                    {
                        for (int i = 1, count = attachments.Count; i <= count; i++)
                        {
                            var fileName = attachments[i].FileName;

                            if (txtAttachments.Text.Length == 0)
                            {
                                txtAttachments.Text = fileName;
                            }
                            else
                            {
                                txtAttachments.AppendText(Environment.NewLine + fileName);
                            }
                        }
                    }
                }
            }
        }

        private void GetCurrentOutlookApplicationInstance()
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
            catch(Exception ex)
            {
                //TODO show exception
                throw new Exception("temp");
            }
        }

        private void InitializeNotifyIcon()
        {
            _notifyIcon = new SuperNotifyIcon();
            _notifyIcon.Icon = Icon;
            _notifyIcon.NotifyIcon.Visible = true;
            _notifyIcon.InitDrop(false);
            _notifyIcon.DragEnter += new DragEventHandler(Icon_DragEnter);
            _notifyIcon.DragDrop += new DragEventHandler(Icon_DragDrop);
        }

        private string GetRecipientSMTPAddress(interop.Recipient recipient)
        {
            // There is an Address property on Recipient, but that can give you an X500 address instead of a 'friendly' email address. Have to dig into COM instead.
            const string PR_SMTP_ADDRESS =
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            interop.PropertyAccessor pa = recipient.PropertyAccessor;
            var property = pa.GetProperty(PR_SMTP_ADDRESS);

            return property != null ? property.ToString() : string.Empty;
        }
    }
}
