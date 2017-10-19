using Microsoft.Exchange.WebServices.Data;
using System;
using System.IO;
using System.Net.Mail;

namespace philre
{
    public static class MailConverter
    {
        public static MailMessage ToMailMessage(this EmailMessage msg)
        {
            // Copy the fieds we need...
            var mailMessage = new MailMessage
            {
                From = msg.From.ToMailAddress(),
                Sender = msg.Sender.ToMailAddress(),
                Subject = msg.Subject,
                Body = msg.Body.Text,
                Priority = msg.Importance.ToMailPriority(),
            };

            if (msg.HasAttachments)
            {
                foreach (var a in msg.Attachments)
                {
                    mailMessage.Attachments.Add(a.ToAttachment());
                }
            }
            foreach (var m in msg.ToRecipients)
            {
                mailMessage.To.Add(m.ToMailAddress());
            }

            foreach (var c in msg.CcRecipients)
            {
                mailMessage.CC.Add(c.ToMailAddress());
            }

            foreach (var b in msg.BccRecipients)
            {
                mailMessage.Bcc.Add(b.ToMailAddress());
            }

            if (msg.InternetMessageHeaders != null)
            {
                foreach (var h in msg.InternetMessageHeaders)
                {
                    mailMessage.Headers.Add(h.Name, h.Value);
                }
            }
            return mailMessage;
        }

        public static System.Net.Mail.Attachment ToAttachment(this Microsoft.Exchange.WebServices.Data.Attachment attachment)
        {
            MemoryStream contentStream = new MemoryStream();
            // Write the attachmentcontent to a MemoryStream
            if (attachment is ItemAttachment)
            {
                var itemAttachment = attachment as ItemAttachment;
                itemAttachment.Load(ItemSchema.MimeContent);
                contentStream = new MemoryStream(itemAttachment.Item.MimeContent.Content);
            }

            if (attachment is FileAttachment)
            {
                var fileAttachment = attachment as FileAttachment;
                fileAttachment.Load(contentStream);
            }

            var result = new System.Net.Mail.Attachment(contentStream, attachment.Name)
            {
                ContentId = attachment.ContentId,
                ContentType = new System.Net.Mime.ContentType(attachment.ContentType),

            };

            result.ContentDisposition.CreationDate = attachment.LastModifiedTime;
            result.ContentDisposition.DispositionType = attachment.ContentType;
            result.ContentDisposition.FileName = attachment.Name;
            result.ContentDisposition.Inline = attachment.IsInline;
            result.ContentDisposition.ModificationDate = attachment.LastModifiedTime;
            result.ContentDisposition.Size = attachment.Size;

            return result;

        }


        public static MailAddress ToMailAddress(this EmailAddress address)
        {
            if (address == null)
            {
                return null;
            }
            return new MailAddress(address.Address, address.Name);
        }

        public static MailPriority ToMailPriority(this Importance importance)
        {

            switch (importance)
            {
                case Importance.High:
                    return MailPriority.High;
                case Importance.Low:
                    return MailPriority.Low;
                default:
                    return MailPriority.Normal;
            }
        }
    }
}
