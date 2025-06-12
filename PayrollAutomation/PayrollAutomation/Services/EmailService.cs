using PayrollAutomation.Models;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Configuration;

namespace PayrollAutomation.Services
{
    internal class EmailService
    {
        private static readonly string apiKey = "YOUR_SENDGRID_API_KEY";
        private static readonly string fromEmail = "your-email@yourdomain.com";
        private static readonly string fromName = "HR Department";

        public static async Task<string> SendEmailAsync(Email email)
        {
            try
            {
                var client = new SendGridClient(apiKey);
                var from = new EmailAddress(fromEmail, fromName);
                var to = new EmailAddress(email.To);
                var msg = MailHelper.CreateSingleEmail(from, to, email.Subject, email.Message, email.Message);

                if (email.HasAttachment && File.Exists(email.AttachmentPath))
                {
                    var fileBytes = File.ReadAllBytes(email.AttachmentPath);
                    var base64Content = Convert.ToBase64String(fileBytes);

                    msg.AddAttachment(Path.GetFileName(email.AttachmentPath), base64Content);
                }

                var response = await client.SendEmailAsync(msg);
                return response.IsSuccessStatusCode ? "Email Sent" : $"Email Failed: {response.StatusCode}";
            }
            catch (Exception ex)
            {
                return $"Email Exception: {ex.Message}";
            }
        }
    }
}
