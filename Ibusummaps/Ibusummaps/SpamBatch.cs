using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.IO;
using System.Net;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;

namespace Ibusummaps
{
    internal class SpamBatch
    {
        #region Constructors

        public SpamBatch(string smtpServer, bool useEncrypted, string username, string password, int messageCount, string bodyFilesPath, string attachmentsPath, string fromAddress, string toAddressListPath,
            string ccAddressListPath = null, string bccAddressStringPath = null)
        {
            if (string.IsNullOrEmpty(smtpServer))
                throw new ArgumentException("The smtp server is required.", smtpServer);

            if (string.IsNullOrEmpty(username)) throw new ArgumentException("Username is required.", username);

            if (string.IsNullOrEmpty(password)) throw new ArgumentException("Password is required.", password);
            
            if (messageCount < 1) throw new ArgumentException("Message count must be greater than 0.", messageCount.ToString());
            MessageCount = messageCount;
            
            if (string.IsNullOrEmpty(bodyFilesPath) || !Directory.Exists(bodyFilesPath))
                throw new ArgumentException("Invalid path to body files.", bodyFilesPath);
            BodyFilesPath = bodyFilesPath;

            if (string.IsNullOrEmpty(attachmentsPath) || !Directory.Exists(attachmentsPath) )
                throw new ArgumentException("Invalid path to attachments.", attachmentsPath);
            AttachmentsPath = attachmentsPath;

            if (string.IsNullOrEmpty(fromAddress))
                throw new ArgumentException("From address is required.", fromAddress);
            //add email address validation
            FromAddress = new MailAddress(fromAddress);

            if (string.IsNullOrEmpty(toAddressListPath))
                throw new ArgumentException("To address file is required.", toAddressListPath);
            if (!File.Exists(toAddressListPath))
                throw new ArgumentException("Invalid path to address list file.", toAddressListPath);
            ToAddressCollection = LoadAddressesFromFile(toAddressListPath);

            if (!string.IsNullOrEmpty(ccAddressListPath) && !File.Exists(ccAddressListPath))
                throw new ArgumentException("Invalid path to cc address list file.", ccAddressListPath);
            CcAddressCollection = ccAddressListPath != null ? LoadAddressesFromFile(ccAddressListPath) : null;


            if (!string.IsNullOrEmpty(bccAddressStringPath) && !File.Exists(bccAddressStringPath))
                throw new ArgumentException("Invalid path to bcc address list file.", bccAddressStringPath);
            BccAddressCollection = bccAddressStringPath != null ? LoadAddressesFromFile(bccAddressStringPath) : null;

            SmtpClient = new SmtpClient(smtpServer)
            {
                //need to prompt for these instead of hardcoding
                Credentials = new NetworkCredential(username, password),
                EnableSsl = useEncrypted
            };

            MessageQueue = new List<MailMessage>();
            QueueMessages(messageCount);
        }

        #endregion

        #region Properties

        private int MessageCount { get; set; }

        private string BodyFilesPath { get; set; }

        private string AttachmentsPath { get; set; }
        
        private MailAddress FromAddress { get; set; }
        
        private MailAddressCollection ToAddressCollection { get; set; }

        private MailAddressCollection CcAddressCollection { get; set; }

        private MailAddressCollection BccAddressCollection { get; set; }

        public List<MailMessage> MessageQueue { get; set; }

        private SmtpClient SmtpClient { get; set; }

        #endregion

        #region Methods

        private static MailAddressCollection LoadAddressesFromFile(string path)
        {
            var retVal = new MailAddressCollection();
            var stream = new StreamReader(path);
            while (!stream.EndOfStream)
            {
                //add email address validation here
                var currentAddy = stream.ReadLine();
                if (currentAddy != null)
                {
                    retVal.Add(currentAddy);
                }
                else
                {
                    throw new WarningException(string.Format("Blank line encountered in file {0}.", path));
                }
            }

            return retVal;
        }

        private static string ChooseRandomFile(string path, int seed)
        {
            var thisDir = new DirectoryInfo(path);
            var availableFiles = thisDir.GetFiles();
            var random = new Random(seed);
            var chosenIndex = random.Next(0, availableFiles.Length - 1);
            return availableFiles[chosenIndex].FullName;
        }

        private void QueueMessages(int messageCount)
        {
            for (var i = 1; i <= messageCount; i++)
            {
                var thisMessage = new MailMessage {From = FromAddress};


                foreach (var toAddress in ToAddressCollection)
                {
                    thisMessage.To.Add(toAddress);
                }
                
                if (CcAddressCollection != null)
                {
                    foreach (var ccAddress in CcAddressCollection)
                    {
                        thisMessage.CC.Add(ccAddress);
                    }
                }

                if (BccAddressCollection != null)
                {
                    foreach (var bccAddress in BccAddressCollection)
                    {
                        thisMessage.Bcc.Add(bccAddress);
                    }                    
                }

                var bodyPath = ChooseRandomFile(BodyFilesPath, i);
                var str = new StreamReader(bodyPath);
                var bodyContent = str.ReadToEnd();
                thisMessage.Body = bodyContent;
                var dirtySubject = bodyContent.Substring(0, 32);
                var rgx = new Regex(@"[\t\n\r]");
                var subject = rgx.Replace(dirtySubject, "");
                thisMessage.Subject = subject;

                var random = new Random();
                var includeAttachment = random.Next(0, 3);
                if (includeAttachment > 0)
                {
                    var attachment = new Attachment(ChooseRandomFile(AttachmentsPath, i));
                    thisMessage.Attachments.Add(attachment);                    
                }

                MessageQueue.Add(thisMessage);
            }
        }

        public void SendNextMessage()
        {
            if (MessageQueue.Count <= 0) return;
            SmtpClient.Send(MessageQueue[0]);
            MessageQueue.RemoveAt(0);
        }

        #endregion
    }
}
