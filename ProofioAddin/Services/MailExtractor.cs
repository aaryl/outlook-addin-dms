using ProofioAddIn.Api;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Services
{
    public sealed class MailExtractor
    {
        private const string PrAttachContentId = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";
        private const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const string PrHtmlBinary = "http://schemas.microsoft.com/mapi/proptag/0x10130102";
        private const string PrBodyUnicode = "http://schemas.microsoft.com/mapi/proptag/0x1000001F";
        private const string PrBodyAnsi = "http://schemas.microsoft.com/mapi/proptag/0x1000001E";

        public EmailPayload Extract(Outlook.MailItem mail, Guid caseId, string forcedDirection = null)
        {
            if (mail == null) throw new ArgumentNullException(nameof(mail));

            var bodyHtml = GetBodyHtml(mail);
            var bodyText = GetBodyText(mail);
            var direction = NormalizeDirection(forcedDirection);

            var payload = new EmailPayload
            {
                CaseId = caseId,
                Subject = SafeString(mail.Subject),
                FromAddr = ResolveSender(mail),
                ToAddrs = ResolveRecipients(mail, Outlook.OlMailRecipientType.olTo),
                CcAddrs = ResolveRecipients(mail, Outlook.OlMailRecipientType.olCC),
                MessageDate = GetMessageDateIso(mail),
                Direction = direction,
                BodyText = bodyText,
                BodyHtml = bodyHtml,
                Attachments = new List<AttachmentPayload>()
            };

            ExtractAttachments(mail, payload);

            if (payload.ToAddrs == null) payload.ToAddrs = new List<string>();
            if (payload.CcAddrs == null) payload.CcAddrs = new List<string>();
            if (payload.Attachments == null) payload.Attachments = new List<AttachmentPayload>();
            if (string.IsNullOrWhiteSpace(payload.MessageDate)) payload.MessageDate = DateTime.UtcNow.ToString("o");
            if (string.IsNullOrWhiteSpace(payload.BodyHtml)) payload.BodyHtml = "<p><em>Kein Inhalt vorhanden.</em></p>";
            if (string.IsNullOrWhiteSpace(payload.BodyText)) payload.BodyText = StripHtmlFallback(payload.BodyHtml);
            if (string.IsNullOrWhiteSpace(payload.Direction)) payload.Direction = "eingehend";

            Logger.Info(
                "Mail extrahiert: Subject='" + payload.Subject +
                "', From='" + payload.FromAddr +
                "', To=" + payload.ToAddrs.Count +
                ", Cc=" + payload.CcAddrs.Count +
                ", Attachments=" + payload.Attachments.Count +
                ", BodyHtmlLength=" + payload.BodyHtml.Length +
                ", BodyTextLength=" + payload.BodyText.Length +
                ", Direction=" + payload.Direction);

            return payload;
        }

        private static string NormalizeDirection(string forcedDirection)
        {
            if (string.Equals(forcedDirection, "ausgehend", StringComparison.OrdinalIgnoreCase))
            {
                return "ausgehend";
            }

            return "eingehend";
        }

        private static string GetMessageDateIso(Outlook.MailItem mail)
        {
            try
            {
                DateTime date;
                if (mail.SentOn != DateTime.MinValue) date = mail.SentOn;
                else if (mail.ReceivedTime != DateTime.MinValue) date = mail.ReceivedTime;
                else if (mail.CreationTime != DateTime.MinValue) date = mail.CreationTime;
                else date = DateTime.Now;

                return date.ToUniversalTime().ToString("o");
            }
            catch
            {
                return DateTime.UtcNow.ToString("o");
            }
        }

        private static string ResolveSender(Outlook.MailItem mail)
        {
            try
            {
                var sender = mail.Sender;
                var smtp = ResolveAddressEntry(sender);
                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(mail.SenderEmailAddress) &&
                    !LooksLikeExchangeLegacyDn(mail.SenderEmailAddress))
                {
                    return mail.SenderEmailAddress;
                }
            }
            catch { }

            try
            {
                var account = mail.SendUsingAccount;
                if (account != null && !string.IsNullOrWhiteSpace(account.SmtpAddress)) return account.SmtpAddress;
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(mail.SentOnBehalfOfName)) return mail.SentOnBehalfOfName;
            }
            catch { }

            return string.Empty;
        }

        private static List<string> ResolveRecipients(Outlook.MailItem mail, Outlook.OlMailRecipientType recipientType)
        {
            var result = new List<string>();
            Outlook.Recipients recipients = null;

            try
            {
                recipients = mail.Recipients;
                if (recipients == null || recipients.Count == 0) return result;

                for (var i = 1; i <= recipients.Count; i++)
                {
                    Outlook.Recipient recipient = null;

                    try
                    {
                        recipient = recipients[i];
                        if (recipient == null || recipient.Type != (int)recipientType) continue;

                        var address = ResolveRecipient(recipient);
                        if (!string.IsNullOrWhiteSpace(address) && !result.Contains(address)) result.Add(address);
                    }
                    finally
                    {
                        if (recipient != null) Marshal.ReleaseComObject(recipient);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Empfänger konnten nicht vollständig aufgelöst werden.", ex);
            }
            finally
            {
                if (recipients != null) Marshal.ReleaseComObject(recipients);
            }

            return result;
        }

        private static string ResolveRecipient(Outlook.Recipient recipient)
        {
            try
            {
                if (!recipient.Resolved) recipient.Resolve();
            }
            catch { }

            try
            {
                var entry = recipient.AddressEntry;
                var smtp = ResolveAddressEntry(entry);
                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(recipient.Address) &&
                    !LooksLikeExchangeLegacyDn(recipient.Address))
                {
                    return recipient.Address;
                }
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(recipient.Name)) return recipient.Name;
            }
            catch { }

            return string.Empty;
        }

        private static string ResolveAddressEntry(Outlook.AddressEntry entry)
        {
            if (entry == null) return string.Empty;

            try
            {
                if (string.Equals(entry.Type, "SMTP", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(entry.Address))
                {
                    return entry.Address;
                }
            }
            catch { }

            try
            {
                var exchangeUser = entry.GetExchangeUser();
                if (exchangeUser != null && !string.IsNullOrWhiteSpace(exchangeUser.PrimarySmtpAddress))
                {
                    return exchangeUser.PrimarySmtpAddress;
                }
            }
            catch { }

            try
            {
                var exchangeDistributionList = entry.GetExchangeDistributionList();
                if (exchangeDistributionList != null &&
                    !string.IsNullOrWhiteSpace(exchangeDistributionList.PrimarySmtpAddress))
                {
                    return exchangeDistributionList.PrimarySmtpAddress;
                }
            }
            catch { }

            try
            {
                var smtp = entry.PropertyAccessor.GetProperty(PrSmtpAddress) as string;
                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(entry.Address) &&
                    !LooksLikeExchangeLegacyDn(entry.Address))
                {
                    return entry.Address;
                }
            }
            catch { }

            return string.Empty;
        }

        private static bool LooksLikeExchangeLegacyDn(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            return value.StartsWith("/O=", StringComparison.OrdinalIgnoreCase) ||
                   value.IndexOf("/OU=", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("/CN=RECIPIENTS/", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string GetBodyHtml(Outlook.MailItem mail)
        {
            var html = TryGetHtmlBody(mail);
            if (!string.IsNullOrWhiteSpace(html)) return html;

            var plain = TryGetPlainBody(mail);
            if (!string.IsNullOrWhiteSpace(plain)) return PlainTextToHtml(plain);

            return "<p><em>Kein Inhalt vorhanden.</em></p>";
        }

        private static string GetBodyText(Outlook.MailItem mail)
        {
            var plain = TryGetPlainBody(mail);
            if (!string.IsNullOrWhiteSpace(plain)) return plain;

            try
            {
                if (!string.IsNullOrWhiteSpace(mail.Body)) return mail.Body;
            }
            catch { }

            return string.Empty;
        }

        private static string TryGetHtmlBody(Outlook.MailItem mail)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(mail.HTMLBody)) return mail.HTMLBody;
            }
            catch { }

            try
            {
                var value = mail.PropertyAccessor.GetProperty(PrHtmlBinary);
                var bytes = value as byte[];
                if (bytes != null && bytes.Length > 0)
                {
                    var html = DecodeHtmlBytes(bytes);
                    if (!string.IsNullOrWhiteSpace(html)) return html;
                }
            }
            catch { }

            return string.Empty;
        }

        private static string TryGetPlainBody(Outlook.MailItem mail)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(mail.Body)) return mail.Body;
            }
            catch { }

            try
            {
                var value = mail.PropertyAccessor.GetProperty(PrBodyUnicode) as string;
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }

            try
            {
                var value = mail.PropertyAccessor.GetProperty(PrBodyAnsi) as string;
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }

            return string.Empty;
        }

        private static string DecodeHtmlBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return string.Empty;

            try
            {
                var utf8 = Encoding.UTF8.GetString(bytes).Trim('\0');
                if (!string.IsNullOrWhiteSpace(utf8)) return utf8;
            }
            catch { }

            try
            {
                var fallback = Encoding.Default.GetString(bytes).Trim('\0');
                if (!string.IsNullOrWhiteSpace(fallback)) return fallback;
            }
            catch { }

            return string.Empty;
        }

        private static string PlainTextToHtml(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "<p><em>Kein Inhalt vorhanden.</em></p>";

            return "<pre style=\"white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;\">" +
                   HtmlEncode(text) +
                   "</pre>";
        }

        private static string StripHtmlFallback(string html)
        {
            if (string.IsNullOrWhiteSpace(html)) return string.Empty;

            var text = html
                .Replace("<br>", "\n")
                .Replace("<br/>", "\n")
                .Replace("<br />", "\n")
                .Replace("</p>", "\n")
                .Replace("</div>", "\n");

            var builder = new StringBuilder(text.Length);
            var insideTag = false;

            foreach (var c in text)
            {
                if (c == '<')
                {
                    insideTag = true;
                    continue;
                }

                if (c == '>')
                {
                    insideTag = false;
                    continue;
                }

                if (!insideTag) builder.Append(c);
            }

            return builder
                .ToString()
                .Replace("&nbsp;", " ")
                .Replace("&amp;", "&")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&quot;", "\"")
                .Replace("&#39;", "'")
                .Trim();
        }

        private static void ExtractAttachments(Outlook.MailItem mail, EmailPayload payload)
        {
            if (payload.Attachments == null) payload.Attachments = new List<AttachmentPayload>();

            Outlook.Attachments attachments = null;

            try
            {
                attachments = mail.Attachments;
                if (attachments == null) return;
                if (attachments.Count == 0) return;

                for (var i = 1; i <= attachments.Count; i++)
                {
                    Outlook.Attachment attachment = null;

                    try
                    {
                        attachment = attachments[i];
                        if (attachment == null) continue;

                        var originalName = string.IsNullOrWhiteSpace(attachment.FileName)
                            ? "attachment.bin"
                            : attachment.FileName;

                        if (IsInlineAttachment(attachment)) continue;

                        var safeName = MakeSafeFileName(originalName);
                        var temp = Path.Combine(
                            Path.GetTempPath(),
                            "Proofio-" + Guid.NewGuid().ToString("N") + "-" + safeName);

                        attachment.SaveAsFile(temp);

                        var bytes = File.ReadAllBytes(temp);
                        TryDelete(temp);

                        payload.Attachments.Add(new AttachmentPayload
                        {
                            Name = originalName,
                            ContentType = GuessContentType(originalName),
                            ContentBase64 = Convert.ToBase64String(bytes)
                        });
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Anhang konnte nicht extrahiert werden.", ex);
                    }
                    finally
                    {
                        if (attachment != null) Marshal.ReleaseComObject(attachment);
                    }
                }
            }
            finally
            {
                if (attachments != null) Marshal.ReleaseComObject(attachments);
            }
        }

        private static bool IsInlineAttachment(Outlook.Attachment attachment)
        {
            try
            {
                var cid = attachment.PropertyAccessor.GetProperty(PrAttachContentId) as string;
                if (string.IsNullOrWhiteSpace(cid)) return false;

                var fileName = attachment.FileName ?? string.Empty;
                var ext = Path.GetExtension(fileName).ToLowerInvariant();

                switch (ext)
                {
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                    case ".gif":
                    case ".bmp":
                    case ".webp":
                    case ".svg":
                        return true;

                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string MakeSafeFileName(string value)
        {
            value = string.IsNullOrWhiteSpace(value) ? "attachment.bin" : value;

            foreach (var c in Path.GetInvalidFileNameChars())
            {
                value = value.Replace(c, '_');
            }

            return value;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        private static string GuessContentType(string fileName)
        {
            var ext = Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant();

            switch (ext)
            {
                case ".xml": return "application/xml";
                case ".pdf": return "application/pdf";
                case ".txt": return "text/plain";
                case ".html":
                case ".htm": return "text/html";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".png": return "image/png";
                case ".gif": return "image/gif";
                case ".doc": return "application/msword";
                case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case ".xls": return "application/vnd.ms-excel";
                case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case ".eml": return "message/rfc822";
                case ".msg": return "application/vnd.ms-outlook";
                case ".csv": return "text/csv";
                case ".zip": return "application/zip";
                default: return "application/octet-stream";
            }
        }

        private static string SafeString(string value)
        {
            return value ?? string.Empty;
        }

        private static string HtmlEncode(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            var builder = new StringBuilder(value.Length);

            foreach (var c in value)
            {
                switch (c)
                {
                    case '&': builder.Append("&amp;"); break;
                    case '<': builder.Append("&lt;"); break;
                    case '>': builder.Append("&gt;"); break;
                    case '"': builder.Append("&quot;"); break;
                    case '\'': builder.Append("&#39;"); break;
                    default: builder.Append(c); break;
                }
            }

            return builder.ToString();
        }
    }
}
