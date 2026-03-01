using System;
using System.Collections.Generic;
using System.IO;

// Mock Exchange Transport Agent Types for Testing
namespace Microsoft.Exchange.Data.Transport
{
    public class SmtpServer
    {
        public string Name { get; set; } = "MockServer";
        public string Version { get; set; } = "15.0.0.0";
    }

    public class RoutingAddress
    {
        private readonly string _address;

        public RoutingAddress(string address)
        {
            _address = address ?? throw new ArgumentNullException(nameof(address));
        }

        public override string ToString() => _address;

        public static implicit operator string(RoutingAddress address) => address?.ToString();
        public static implicit operator RoutingAddress(string address) => new RoutingAddress(address);
    }

    public class MailItem
    {
        public RoutingAddress FromAddress { get; set; }
        public Email.EmailMessage Message { get; set; }
        public string MessageId => Message?.MessageId ?? Guid.NewGuid().ToString();
        public List<RoutingAddress> Recipients { get; set; } = new List<RoutingAddress>();
        public Dictionary<string, object> Properties { get; set; } = new Dictionary<string, object>();
    }
}

namespace Microsoft.Exchange.Data.Transport.Routing
{
    public abstract class RoutingAgentFactory
    {
        public abstract RoutingAgent CreateAgent(SmtpServer server);
    }

    public abstract class RoutingAgent
    {
        public event EventHandler<QueuedMessageEventArgs> OnSubmittedMessage;
        public event EventHandler<QueuedMessageEventArgs> OnResolvedMessage;

        protected void InvokeOnSubmittedMessage(QueuedMessageEventArgs args)
        {
            OnSubmittedMessage?.Invoke(new SubmittedMessageEventSource(), args);
        }

        protected void InvokeOnResolvedMessage(QueuedMessageEventArgs args)
        {
            OnResolvedMessage?.Invoke(new ResolvedMessageEventSource(), args);
        }
    }

    public class QueuedMessageEventArgs : EventArgs
    {
        public MailItem MailItem { get; set; }
    }

    public class SubmittedMessageEventSource { }
    public class ResolvedMessageEventSource { }
}

namespace Microsoft.Exchange.Data.Transport.Email
{
    public enum BodyFormat
    {
        Text,
        Html
    }

    public class Body
    {
        private string _text;
        private BodyFormat _format;

        public Body(string text, BodyFormat format = BodyFormat.Text)
        {
            _text = text ?? string.Empty;
            _format = format;
        }

        public string GetText(BodyFormat format = BodyFormat.Text)
        {
            if (format == _format)
                return _text;
            
            // Simple conversion for testing
            if (format == BodyFormat.Html && _format == BodyFormat.Text)
                return $"<html><body><pre>{_text}</pre></body></html>";
            
            if (format == BodyFormat.Text && _format == BodyFormat.Html)
                return System.Text.RegularExpressions.Regex.Replace(_text, "<[^>]*>", "");
                
            return _text;
        }

        public void SetText(string text)
        {
            _text = text ?? string.Empty;
        }
    }

    public class EmailMessage
    {
        public string MessageId { get; set; } = $"<{Guid.NewGuid()}@mock.server>";
        public Body Body { get; set; }
        public AttachmentCollection Attachments { get; set; } = new AttachmentCollection();
        public MimeDocument MimeDocument { get; set; } = new MimeDocument();
        public string Subject { get; set; } = string.Empty;
        public RoutingAddress From { get; set; }
    }

    public class AttachmentCollection : List<Attachment>
    {
        // Implements ICollection interface for attachments
    }

    public class Attachment
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public Stream ContentStream { get; set; }
        public long Size { get; set; }
    }

    public class MimeDocument
    {
        public MimePart RootPart { get; set; } = new MimePart();
    }

    public class MimePart
    {
        public ContentTypeHeader ContentType { get; set; } = new ContentTypeHeader();
    }

    public class ContentTypeHeader
    {
        public string MediaType { get; set; } = "text/plain";
        public string SubType { get; set; } = "plain";
        public Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();

        public override string ToString()
        {
            var result = $"{MediaType}/{SubType}";
            foreach (var param in Parameters)
            {
                result += $"; {param.Key}={param.Value}";
            }
            return result;
        }
    }
}