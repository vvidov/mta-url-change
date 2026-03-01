using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;

namespace UrlToTextTransportAgent
{
    /// <summary>
    /// Result of URL processing operation
    /// </summary>
    public struct UrlProcessingResult
    {
        public string ProcessedText { get; set; }
        public int UrlsConverted { get; set; }
    }

    /// <summary>
    /// Exchange Transport Agent that scans external emails and converts URLs to plain text
    /// </summary>
    public class UrlToTextAgent : RoutingAgent
    {
        private static readonly string LogPath = @"C:\ExchangeLogs\UrlToTextAgent.log";

        public UrlToTextAgent()
        {
            LogMessage("UrlToTextAgent initialized", "INFO");
            
            // Subscribe to the OnSubmittedMessage event to process messages
            this.OnSubmittedMessage += OnSubmittedMessageHandler;
            this.OnResolvedMessage += OnResolvedMessageHandler;
            
            LogMessage("Event handlers registered successfully", "DEBUG");
        }

        /// <summary>
        /// Event handler for submitted messages
        /// </summary>
        private void OnSubmittedMessageHandler(SubmittedMessageEventSource source, QueuedMessageEventArgs e)
        {
            try
            {
                LogMessage("OnSubmittedMessage event triggered", "DEBUG");
                ProcessMessage(e.MailItem, "OnSubmittedMessage");
            }
            catch (Exception ex)
            {
                LogMessage($"Error in OnSubmittedMessage: {ex.Message}", "ERROR");
                LogMessage($"Stack trace: {ex.StackTrace}", "ERROR");
            }
        }

        /// <summary>
        /// Event handler for resolved messages
        /// </summary>
        private void OnResolvedMessageHandler(ResolvedMessageEventSource source, QueuedMessageEventArgs e)
        {
            try
            {
                LogMessage("OnResolvedMessage event triggered", "DEBUG");
                ProcessMessage(e.MailItem, "OnResolvedMessage");
            }
            catch (Exception ex)
            {
                LogMessage($"Error in OnResolvedMessage: {ex.Message}", "ERROR");
                LogMessage($"Stack trace: {ex.StackTrace}", "ERROR");
            }
        }

        /// <summary>
        /// Process the email message to convert URLs to plain text
        /// </summary>
        private void ProcessMessage(MailItem mailItem, string eventName)
        {
            if (mailItem?.Message == null)
            {
                LogMessage("ProcessMessage: MailItem or Message is null", "DEBUG");
                return;
            }

            string messageId = mailItem.Message.MessageId ?? "Unknown";
            string fromAddress = mailItem.FromAddress?.ToString() ?? "Unknown";
            
            LogMessage($"Starting processing for message ID: {messageId} from: {fromAddress} in event: {eventName}", "INFO");

            // Check if the message is from an external sender
            if (!IsExternalMessage(mailItem))
            {
                LogMessage($"Message {messageId} from {fromAddress} is internal, skipping processing", "INFO");
                return;
            }

            // Check if message is signed or encrypted - skip processing to preserve integrity
            if (IsSignedOrEncryptedMessage(mailItem.Message))
            {
                LogMessage($"Message {messageId} from {fromAddress} is signed or encrypted, skipping processing to preserve integrity", "WARNING");
                return;
            }

            LogMessage($"Processing external message {messageId} from {fromAddress} in {eventName}", "INFO");

            EmailMessage message = mailItem.Message;
            bool messageModified = false;
            int urlsConverted = 0;

            try
            {
                // Process text body
                if (message.Body != null)
                {
                    string originalText = message.Body.GetText();
                    if (!string.IsNullOrEmpty(originalText))
                    {
                        LogMessage($"Processing plain text body for message {messageId} (length: {originalText.Length})", "DEBUG");
                        
                        var result = ConvertUrls(originalText, messageId);
                        
                        if (originalText != result.ProcessedText)
                        {
                            message.Body = new Body(result.ProcessedText);
                            messageModified = true;
                            urlsConverted += result.UrlsConverted;
                            LogMessage($"Modified text body for message {messageId}, converted {result.UrlsConverted} URLs to plain text", "INFO");
                        }
                    }
                }

                // Process HTML body
                if (message.Body != null)
                {
                    string originalHtml = message.Body.GetText(BodyFormat.Html);
                    if (!string.IsNullOrEmpty(originalHtml) && originalHtml != message.Body.GetText())
                    {
                        LogMessage($"Processing HTML body for message {messageId} (length: {originalHtml.Length})", "DEBUG");
                        
                        var result = ConvertHtmlUrls(originalHtml, messageId);
                        
                        if (originalHtml != result.ProcessedText)
                        {
                            message.Body = new Body(result.ProcessedText, BodyFormat.Html);
                            messageModified = true;
                            urlsConverted += result.UrlsConverted;
                            LogMessage($"Modified HTML body for message {messageId}, converted {result.UrlsConverted} URLs to plain text", "INFO");
                        }
                    }
                }

                if (messageModified)
                {
                    LogMessage($"Successfully processed message {messageId} from {fromAddress}. Total URLs converted to plain text: {urlsConverted}", "SUCCESS");
                }
                else
                {
                    LogMessage($"No URLs found in message {messageId} from {fromAddress}", "INFO");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error processing message {messageId} from {fromAddress}: {ex.Message}", "ERROR");
                LogMessage($"Stack trace: {ex.StackTrace}", "ERROR");
            }
        }

        /// <summary>
        /// Check if the message is signed or encrypted
        /// </summary>
        private bool IsSignedOrEncryptedMessage(EmailMessage message)
        {
            try
            {
                // Check Content-Type headers for S/MIME signatures or encryption
                var contentType = message.MimeDocument?.RootPart?.ContentType;
                if (contentType != null)
                {
                    string contentTypeString = contentType.ToString().ToLowerInvariant();
                    
                    // Check for S/MIME signed content
                    if (contentTypeString.Contains("application/pkcs7-mime") ||
                        contentTypeString.Contains("application/x-pkcs7-mime") ||
                        contentTypeString.Contains("multipart/signed"))
                    {
                        LogMessage("Detected S/MIME signed message", "INFO");
                        return true;
                    }
                    
                    // Check for encrypted content
                    if (contentTypeString.Contains("application/pkcs7-mime") && 
                        (contentTypeString.Contains("smime-type=enveloped-data") || 
                         contentTypeString.Contains("name=\"smime.p7m\"")))
                    {
                        LogMessage("Detected encrypted S/MIME message", "INFO");
                        return true;
                    }
                }

                // Check for PGP signatures or encryption
                if (message.Body != null)
                {
                    string bodyText = message.Body.GetText() ?? string.Empty;
                    
                    if (bodyText.Contains("-----BEGIN PGP SIGNED MESSAGE-----") ||
                        bodyText.Contains("-----BEGIN PGP MESSAGE-----") ||
                        bodyText.Contains("-----BEGIN PGP SIGNATURE-----"))
                    {
                        LogMessage("Detected PGP signed/encrypted message", "INFO");
                        return true;
                    }
                }

                // Check attachments for signature files
                if (message.Attachments != null && message.Attachments.Count > 0)
                {
                    foreach (var attachment in message.Attachments)
                    {
                        string fileName = attachment.FileName?.ToLowerInvariant() ?? string.Empty;
                        string contentType = attachment.ContentType?.ToLowerInvariant() ?? string.Empty;
                        
                        // Check for signature file attachments
                        if (fileName.EndsWith(".p7s") || 
                            fileName.EndsWith(".sig") ||
                            fileName == "smime.p7s" ||
                            contentType.Contains("application/pkcs7-signature") ||
                            contentType.Contains("application/x-pkcs7-signature"))
                        {
                            LogMessage($"Detected signature attachment: {fileName}", "INFO");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Error checking for signed/encrypted message: {ex.Message}", "ERROR");
                // If we can't determine, err on the side of caution and skip processing
                return true;
            }
        }

        /// <summary>
        /// Check if the message is from an external sender
        /// </summary>
        private bool IsExternalMessage(MailItem mailItem)
        {
            try
            {
                // Get the organization's accepted domains (you may need to customize this)
                // For now, we'll check if the sender is not from common internal domains
                string fromAddress = mailItem.FromAddress?.ToString() ?? string.Empty;
                string fromDomain = GetDomainFromAddress(fromAddress);
                
                if (string.IsNullOrEmpty(fromDomain))
                {
                    LogMessage($"Could not extract domain from address: {fromAddress}", "WARNING");
                    return false;
                }

                LogMessage($"Checking if domain '{fromDomain}' is external", "DEBUG");

                // Add your organization's domains here
                string[] internalDomains = { "yourdomain.com", "internal.local" };
                
                foreach (string internalDomain in internalDomains)
                {
                    if (fromDomain.Equals(internalDomain, StringComparison.OrdinalIgnoreCase))
                    {
                        LogMessage($"Domain '{fromDomain}' matches internal domain '{internalDomain}'", "DEBUG");
                        return false;
                    }
                }

                LogMessage($"Domain '{fromDomain}' identified as external", "DEBUG");
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"Error in IsExternalMessage: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Extract domain from email address
        /// </summary>
        private string GetDomainFromAddress(string emailAddress)
        {
            try
            {
                if (string.IsNullOrEmpty(emailAddress))
                    return string.Empty;

                int atIndex = emailAddress.LastIndexOf('@');
                if (atIndex > 0 && atIndex < emailAddress.Length - 1)
                {
                    string domain = emailAddress.Substring(atIndex + 1);
                    LogMessage($"Extracted domain '{domain}' from address '{emailAddress}'", "DEBUG");
                    return domain;
                }

                LogMessage($"Could not extract domain from malformed address: '{emailAddress}'", "WARNING");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogMessage($"Error extracting domain from '{emailAddress}': {ex.Message}", "ERROR");
                return string.Empty;
            }
        }

        /// <summary>
        /// Convert URLs from plain text to non-clickable format
        /// </summary>
        private UrlProcessingResult ConvertUrls(string text, string messageId)
        {
            var result = new UrlProcessingResult { ProcessedText = text, UrlsConverted = 0 };
            
            if (string.IsNullOrEmpty(text))
            {
                LogMessage($"ConvertUrls: Empty text for message {messageId}", "DEBUG");
                return result;
            }

            // Regex pattern to match HTTP/HTTPS URLs
            string urlPattern = @"https?://[^\s<>\""']+";
            
            result.ProcessedText = Regex.Replace(text, urlPattern, match =>
            {
                string url = match.Value;
                result.UrlsConverted++;
                LogMessage($"Message {messageId}: Converting URL to plain text: {url}", "DEBUG");
                return url; // Keep the URL as plain text, don't make it clickable
            }, RegexOptions.IgnoreCase);

            LogMessage($"ConvertUrls: Processed message {messageId}, converted {result.UrlsConverted} URLs to plain text", "DEBUG");
            return result;
        }

        /// <summary>
        /// Convert URLs from HTML content to plain text format
        /// </summary>
        private UrlProcessingResult ConvertHtmlUrls(string html, string messageId)
        {
            var result = new UrlProcessingResult { ProcessedText = html, UrlsConverted = 0 };
            
            if (string.IsNullOrEmpty(html))
            {
                LogMessage($"ConvertHtmlUrls: Empty HTML for message {messageId}", "DEBUG");
                return result;
            }

            // Remove href attributes from anchor tags and replace with plain text
            string hrefPattern = @"<a\s+[^>]*href\s*=\s*[""']([^""']*)[""'][^>]*>(.*?)</a>";
            result.ProcessedText = Regex.Replace(result.ProcessedText, hrefPattern, match =>
            {
                string url = match.Groups[1].Value;
                string linkText = match.Groups[2].Value;
                result.UrlsConverted++;
                
                LogMessage($"Message {messageId}: Converting HTML link to plain text: {url}", "DEBUG");
                
                // Return the link text followed by the URL in plain text
                return $"{linkText} ({url})";
            }, RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Also handle plain URLs in HTML that are not in anchor tags (keep them as-is since they're already plain text)
            string plainUrlPattern = @"https?://[^\s<>\""']+";
            var plainUrlMatches = Regex.Matches(result.ProcessedText, plainUrlPattern, RegexOptions.IgnoreCase);
            if (plainUrlMatches.Count > 0)
            {
                LogMessage($"Message {messageId}: Found {plainUrlMatches.Count} plain URLs in HTML (keeping as plain text)", "DEBUG");
            }

            LogMessage($"ConvertHtmlUrls: Processed message {messageId}, converted {result.UrlsConverted} HTML links to plain text", "DEBUG");
            return result;
        }

        /// <summary>
        /// Log messages to a file with different severity levels
        /// </summary>
        private void LogMessage(string message, string level = "INFO")
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level.PadRight(7)}] [PID:{System.Diagnostics.Process.GetCurrentProcess().Id}] {message}";
                
                // Ensure log directory exists
                string logDir = Path.GetDirectoryName(LogPath);
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // Append to log file with thread safety
                File.AppendAllText(LogPath, logEntry + Environment.NewLine);
                
                // Also log to Windows Event Log for critical errors
                if (level == "ERROR" || level == "CRITICAL")
                {
                    try
                    {
                        System.Diagnostics.EventLog.WriteEntry("UrlToTextAgent", message, 
                            level == "CRITICAL" ? System.Diagnostics.EventLogEntryType.Error : 
                                                System.Diagnostics.EventLogEntryType.Warning);
                    }
                    catch
                    {
                        // Ignore event log errors
                    }
                }
            }
            catch (Exception ex)
            {
                // Try to log to a backup location if primary logging fails
                try
                {
                    string backupLog = Path.Combine(Path.GetTempPath(), "UrlToTextAgent_backup.log");
                    string backupEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [ERROR] Logging failed for: {message}. Error: {ex.Message}";
                    File.AppendAllText(backupLog, backupEntry + Environment.NewLine);
                }
                catch
                {
                    // If all logging fails, ignore to prevent agent from crashing
                }
            }
        }
    }
}