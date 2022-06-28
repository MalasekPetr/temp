using Microsoft.AspNetCore.Http;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;

namespace NBS.MailBox.BackEnd.Models
{
    public class Message
    {
        public Message()
        {

        }

        public Message(MimeMessage source)
        {
            Importance = source.Importance;
            Priority = source.Priority;
            if (source.Sender != null)
            {
                Sender = source.Sender.ToString();
            }
            From = ((MailAddressCollection)source.From).ToString();
            To = ((MailAddressCollection)source.To).ToString();
            Cc = ((MailAddressCollection)source.Cc).ToString();
            Bcc = ((MailAddressCollection)source.Bcc).ToString();
            Subject = source.Subject;
            Date = source.Date;
            References = source.References;
            InReplyTo = source.InReplyTo;
            MessageId = source.MessageId;
            MimeVersion = source.MimeVersion;
            TextBody = source.TextBody;
            HtmlBody = source.HtmlBody;
            if (source.Headers != null)
            {
                Headers = new HeaderList();
                foreach (var header in source.Headers)
                {
                    Headers.Add(header);
                }
            }
            if (source.Attachments != null)
            {
                // Attachments = new List<IFormFile>();
                Files = new List<string>();
                foreach (var attachment in source.Attachments)
                {
                    if (attachment is MessagePart)
                    {
                        var fileName = attachment.ContentDisposition?.FileName;
                        // var rfc822 = att;

                        if (string.IsNullOrEmpty(fileName))
                            fileName = "attached-message.eml";

                        // using var ms = new MemoryStream();
                        // rfc822.Message.WriteTo(ms);
                        // Attachments.Add(new FormFile(ms, 0, ms.Length, fileName, fileName));
                        Files.Add(fileName);
                    }
                    else
                    {
                        var part = (MimePart)attachment;
                        var fileName = part.FileName;

                        // using var ms = new MemoryStream();
                        // part.Content.DecodeTo(ms);
                        // Attachments.Add(new FormFile(ms, 0, ms.Length, fileName, fileName));
                        Files.Add(fileName);
                    }
                }
            }
        }

        public MessageImportance Importance { get; set; }
        public MessagePriority Priority { get; set; }
        public string Sender { get; set; }
        public string From { get; set; }
        public string ReplyTo { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        public string Subject { get; set; }
        public DateTimeOffset Date { get; set; }
        public MessageIdList References { get; set; }
        public string InReplyTo { get; set; }
        public string MessageId { get; set; }
        public Version MimeVersion { get; set; }
        public string TextBody { get; set; }
        public string HtmlBody { get; set; }
        public HeaderList Headers { get; set; }
        public List<IFormFile> Attachments { get; set; }
        public List<string> Files { get; set; }
    }
}