using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ibusummaps
{
    class Program
    {
        static void Main(string[] args)
        {
            var smtpServer = args[0];
            var useEncrypted = bool.Parse(args[1]);
            var username = args[2];
            var password = args[3];
            var messageCount = int.Parse(args[4]);
            var bodyFilesPath = args[5];
            var attachmentsPath = args[6];
            var fromAddress = args[7];
            var toAddressListPath = args[8];
            var ccAddressListPath = (args.Length > 9) ? args[9] : null;
            var bccAddressListPath = (args.Length > 10) ? args[10] : null;

            var spamBatch = new SpamBatch(smtpServer, useEncrypted, username, password, messageCount, bodyFilesPath, attachmentsPath,
                fromAddress, toAddressListPath, ccAddressListPath, bccAddressListPath);

            while (spamBatch.MessageQueue.Count > 0)
            {
                spamBatch.SendNextMessage();
                System.Threading.Thread.Sleep(1000);
            }
        }
    }
}
