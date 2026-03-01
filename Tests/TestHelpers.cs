using System;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;

namespace UrlToTextTransportAgent.Tests
{
    public static class TestHelpers
    {
        public static MailItem CreateTestMailItem(string fromAddress, string messageBody, BodyFormat bodyFormat = BodyFormat.Text)
        {
            var mailItem = new MailItem
            {
                FromAddress = new RoutingAddress(fromAddress),
                Message = new EmailMessage
                {
                    MessageId = $"<test-{Guid.NewGuid()}@test.com>",
                    Body = new Body(messageBody, bodyFormat),
                    Subject = "Test Message"
                }
            };

            return mailItem;
        }

        public static MailItem CreateSignedTestMailItem(string fromAddress, string messageBody)
        {
            var mailItem = CreateTestMailItem(fromAddress, messageBody);
            
            // Simulate signed message by setting content type
            mailItem.Message.MimeDocument.RootPart.ContentType.MediaType = "application";
            mailItem.Message.MimeDocument.RootPart.ContentType.SubType = "pkcs7-mime";
            mailItem.Message.MimeDocument.RootPart.ContentType.Parameters["smime-type"] = "signed-data";

            return mailItem;
        }

        public static MailItem CreateEncryptedTestMailItem(string fromAddress, string messageBody)
        {
            var mailItem = CreateTestMailItem(fromAddress, messageBody);
            
            // Simulate encrypted message by setting content type
            mailItem.Message.MimeDocument.RootPart.ContentType.MediaType = "application";
            mailItem.Message.MimeDocument.RootPart.ContentType.SubType = "pkcs7-mime";
            mailItem.Message.MimeDocument.RootPart.ContentType.Parameters["smime-type"] = "enveloped-data";

            return mailItem;
        }

        public static MailItem CreatePgpSignedTestMailItem(string fromAddress, string messageBody)
        {
            var pgpMessage = @"-----BEGIN PGP SIGNED MESSAGE-----
Hash: SHA256

" + messageBody + @"

-----BEGIN PGP SIGNATURE-----
iQEcBAEBCAAGBQJhxtest...
-----END PGP SIGNATURE-----";

            return CreateTestMailItem(fromAddress, pgpMessage);
        }

        public static void AddSignatureAttachment(MailItem mailItem)
        {
            var attachment = new Attachment
            {
                FileName = "smime.p7s",
                ContentType = "application/pkcs7-signature"
            };
            
            mailItem.Message.Attachments.Add(attachment);
        }

        public static string CreateHtmlWithLinks(string baseText, params string[] urls)
        {
            var html = $"<html><body><p>{baseText}</p>";
            
            foreach (var url in urls)
            {
                html += $"<p><a href=\"{url}\">Visit {url}</a></p>";
            }
            
            html += "</body></html>";
            return html;
        }

        public static string CreatePlainTextWithUrls(string baseText, params string[] urls)
        {
            var text = baseText + Environment.NewLine;
            
            foreach (var url in urls)
            {
                text += $"Check out: {url}" + Environment.NewLine;
            }
            
            return text;
        }
    }
}