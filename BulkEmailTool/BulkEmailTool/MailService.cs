using System;
using System.Net;
using System.Net.Mail;

namespace BulkEmailTool
{
    public class MailService
    {
        public MailService()
        {

        }

        public void SendMail(Promoter promoter)
        {
            Console.WriteLine($"Sending mail to ${promoter.Name} at ${promoter.EmailAddress}");

            try
            {
                using (var message = new MailMessage())
                {
                    message.To.Add(new MailAddress(promoter.EmailAddress, promoter.Name));
                    message.From = new MailAddress(AppSettings.EmailAddres, "Hard Fact");
                    message.Subject = "Subject";
                    message.Body = "Body";
                    message.IsBodyHtml = true;

                    using (var client = new SmtpClient("smtp.gmail.com"))
                    {
                        client.Port = 587;
                        client.Credentials = new NetworkCredential(AppSettings.EmailAddres, AppSettings.EmailPassword);
                        client.EnableSsl = true;
                        client.Send(message);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error sending mail to {promoter.EmailAddress}. {e}");
            }

        }
    }
}