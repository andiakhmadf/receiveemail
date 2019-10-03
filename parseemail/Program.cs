using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EAGetMail;

namespace parseemail
{
    class Program
    {
        static void ParseEmail(string emlFile)
        {
            Mail oMail = new Mail("TryIt");
            oMail.Load(emlFile, false);

            // Parse Mail From, Sender
            Console.WriteLine("From: {0}", oMail.From.ToString());

            // Parse Mail To, Recipient
            MailAddress[] addrs = oMail.To;
            for (int i = 0; i < addrs.Length; i++)
            {
                Console.WriteLine("To: {0}", addrs[i].ToString());
            }

            // Parse Mail CC
            addrs = oMail.Cc;
            for (int i = 0; i < addrs.Length; i++)
            {
                Console.WriteLine("To: {0}", addrs[i].ToString());
            }

            // Parse Mail Subject
            //Console.WriteLine("Subject: {0}", oMail.Subject);

            // Parse Mail Text/Plain body
            Console.WriteLine("TextBody: {0}", oMail.TextBody);

            // Parse Mail Html Body
            Console.WriteLine("HtmlBody: {0}", oMail.HtmlBody);

            // Parse Attachments
            Attachment[] atts = oMail.Attachments;
            for (int i = 0; i < atts.Length; i++)
            {
                Console.WriteLine("Attachment: {0}", atts[i].Name);
            }
        }

        static void Main(string[] args)
        {
            try
            {
                ParseEmail("D:\\test2.eml");
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }
        }
    }
}
