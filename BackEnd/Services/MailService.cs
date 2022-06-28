using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;
using MimeKit;
using NBS.MailBox.BackEnd.Models;
using System.IO;
using System.Threading.Tasks;

namespace NBS.MailBox.BackEnd.Services
{
    public class MailService : IMailService
    {
        private readonly MailSettings _mailSettings;
        public MailService(IOptions<MailSettings> mailSettings)
        {
            _mailSettings = mailSettings.Value;
        }

        public async Task SendEmailAsync(Message message)
        {
            var mimemessage = new MimeMessage
            {
                Sender = MailboxAddress.Parse(message.Sender)
            };
            mimemessage.To.Add(MailboxAddress.Parse(message.To));
            if (!string.IsNullOrEmpty(message.Cc))
            {
                mimemessage.Cc.Add(MailboxAddress.Parse(message.Cc));
            }
            if (!string.IsNullOrEmpty(message.Bcc))
            {
                mimemessage.Cc.Add(MailboxAddress.Parse(message.Bcc));
            }
            mimemessage.Subject = message.Subject;
            var builder = new BodyBuilder();
            if (message.Attachments != null)
            {
                byte[] fileBytes;
                foreach (var file in message.Attachments)
                {
                    if (file.Length > 0)
                    {
                        using (var ms = new MemoryStream())
                        {
                            file.CopyTo(ms);
                            fileBytes = ms.ToArray();
                        }
                        builder.Attachments.Add(file.FileName, fileBytes, ContentType.Parse(file.ContentType));
                    }
                }
            }
            mimemessage.InReplyTo = message.InReplyTo;
            mimemessage.MessageId = message.MessageId;
            mimemessage.Priority = message.Priority;
            mimemessage.Date = message.Date;
            builder.HtmlBody = message.HtmlBody;
            mimemessage.Body = builder.ToMessageBody();
            using var smtp = new SmtpClient();
            smtp.Connect(_mailSettings.Host, _mailSettings.Port, SecureSocketOptions.StartTls);
            await smtp.SendAsync(mimemessage);
            smtp.Disconnect(true);
        }
    }
}
