using System;
using Microsoft.Office.Interop.Outlook;

namespace Outlook_mail_reader
{
    class Program
    {
        public static Application myApp;

        public static void Main(string[] args)
        {
            ReadMail();
            Console.ReadKey();
        }

        static void ReadMail()
        {
            Application app = null;
            _NameSpace ns = null;
            MAPIFolder inboxFolder = null;


            app = new Application();
            ns = app.GetNamespace("MAPI");

            inboxFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            
            Console.WriteLine("Folder Name: {0}, EntryId: {1}", inboxFolder.Name, inboxFolder.EntryID);
            Console.WriteLine("Num Items: {0}", inboxFolder.Items.Count.ToString());
            string subject = "type your subject line or partial subject line which you want to search?";
            string filter = $"@SQL=\"urn:schemas:mailheader:subject\" like\'%{subject}%\'";
            var searchResult = inboxFolder.Items.Find(filter); 

            Console.WriteLine("Subject: {0}", searchResult.Subject);
            Console.WriteLine("Sent: {0} {1}", searchResult.SentOn.ToLongDateString(), searchResult.SentOn.ToLongTimeString());
            Console.WriteLine("Sendername: {0}", searchResult.SenderName);
            Console.WriteLine("Body: {0}", searchResult.Body);

            Console.WriteLine("Do you want to print all mails? Y/N");
            if(Console.ReadLine().ToLower() == "y")
            {
                PrintAllMails(inboxFolder);
            }
        }

        static void PrintAllMails(MAPIFolder inboxFolder)
        {
            for (int counter = 1; counter <= inboxFolder.Items.Count; counter++)
            {
                Console.Write(inboxFolder.Items.Count + " " + counter);
                dynamic item = inboxFolder.Items[counter];
                Console.WriteLine("Item: {0}", counter.ToString());
                Console.WriteLine("Subject: {0}", item.Subject);
                Console.WriteLine("Sent: {0} {1}", item.SentOn.ToLongDateString(), item.SentOn.ToLongTimeString());
                Console.WriteLine("Sendername: {0}", item.SenderName);
                Console.WriteLine("Body: {0}", item.Body);
            }
        }
    }
}
