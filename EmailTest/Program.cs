using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;

namespace EmailTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(Directory.GetCurrentDirectory());
            SendEmail();
            Console.WriteLine("Done");
            Console.ReadLine();
        }

        public static void SendEmail()
        {
            try
            {
                var outlook = new Application();

                var email = (MailItem)outlook.CreateItem(OlItemType.olMailItem);

                email.Body = "This is a programmatically generated email.";

                var attachments = new string[] { "Program.cs", "Another Blue Abstract Desktop.jpg" };
                
                foreach(var file in attachments)
                {
                    var attachment = email.Attachments.Add(Path.Combine(Directory.GetCurrentDirectory(), file), OlAttachmentType.olByValue, email.Body.Length + 1, file);
                }

                email.Subject = "Test email";

                var recipient = email.Recipients.Add("waters.daniel.c@gmail.com");
                recipient.Resolve();

                email.Send();
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}
