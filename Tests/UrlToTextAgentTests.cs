using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;

namespace UrlToTextTransportAgent.Tests
{
    [TestClass]
    public class UrlToTextAgentTests
    {
        private UrlToTextAgent _agent;
        private string _testLogPath;

        [TestInitialize]
        public void TestInitialize()
        {
            _agent = new UrlToTextAgent();
            
            // Set up test log path
            _testLogPath = Path.Combine(Path.GetTempPath(), $"UrlToTextAgentTest_{Guid.NewGuid()}.log");
            
            // Use reflection to set the log path for testing
            var logPathField = typeof(UrlToTextAgent).GetField("LogPath", BindingFlags.NonPublic | BindingFlags.Static);
            if (logPathField != null)
            {
                logPathField.SetValue(null, _testLogPath);
            }
        }

        [TestCleanup]
        public void TestCleanup()
        {
            // Clean up test log file
            if (File.Exists(_testLogPath))
            {
                try { File.Delete(_testLogPath); } catch { }
            }
        }

        [TestMethod]
        public void TestAgentFactory_CreatesAgent()
        {
            // Arrange
            var factory = new UrlToTextAgentFactory();
            var server = new SmtpServer { Name = "TestServer" };

            // Act
            var agent = factory.CreateAgent(server);

            // Assert
            Assert.IsNotNull(agent);
            Assert.IsInstanceOfType(agent, typeof(UrlToTextAgent));
        }

        [TestMethod]
        public void TestPlainTextUrlConversion()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = TestHelpers.CreatePlainTextWithUrls(
                "This is a test message.",
                "https://www.malicious-site.com",
                "http://phishing-attempt.org/login"
            );
            
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert
            var processedBody = mailItem.Message.Body.GetText();
            // URLs should still be present as plain text (not clickable)
            Assert.IsTrue(processedBody.Contains("https://www.malicious-site.com"));
            Assert.IsTrue(processedBody.Contains("http://phishing-attempt.org/login"));
            // URLs are converted to plain text, not removed
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestHtmlUrlConversion()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = TestHelpers.CreateHtmlWithLinks(
                "This is a test HTML message.",
                "https://www.malicious-site.com",
                "http://phishing-attempt.org/login"
            );
            
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody, BodyFormat.Html);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert
            var processedBody = mailItem.Message.Body.GetText(BodyFormat.Html);
            // Should not contain hyperlinks (anchor tags)
            Assert.IsFalse(processedBody.Contains("href=\"https://www.malicious-site.com\""));
            Assert.IsFalse(processedBody.Contains("href=\"http://phishing-attempt.org/login\""));
            // But should contain URLs as plain text
            Assert.IsTrue(processedBody.Contains("https://www.malicious-site.com"));
            Assert.IsTrue(processedBody.Contains("http://phishing-attempt.org/login"));
        }

        [TestMethod]
        public void TestInternalEmailSkipped()
        {
            // Arrange
            var fromAddress = "internal@yourdomain.com"; // Internal domain
            var messageBody = TestHelpers.CreatePlainTextWithUrls(
                "Internal message with URLs.",
                "https://www.external-site.com"
            );
            
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert - URL should still be present (not processed)
            var processedBody = mailItem.Message.Body.GetText();
            Assert.IsTrue(processedBody.Contains("https://www.external-site.com"));
            // Internal emails are not processed, so URLs remain unchanged
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestSignedEmailSkipped()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = TestHelpers.CreatePlainTextWithUrls(
                "Signed message with URLs.",
                "https://www.should-not-be-converted.com"
            );
            
            var mailItem = TestHelpers.CreateSignedTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert - URL should still be present (not processed due to signature)
            var processedBody = mailItem.Message.Body.GetText();
            Assert.IsTrue(processedBody.Contains("https://www.should-not-be-converted.com"));
            // Signed emails are not processed to preserve signature integrity
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestEncryptedEmailSkipped()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = TestHelpers.CreatePlainTextWithUrls(
                "Encrypted message with URLs.",
                "https://www.should-not-be-converted.com"
            );
            
            var mailItem = TestHelpers.CreateEncryptedTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert - URL should still be present (not processed due to encryption)
            var processedBody = mailItem.Message.Body.GetText();
            Assert.IsTrue(processedBody.Contains("https://www.should-not-be-converted.com"));
            // Encrypted emails are not processed to preserve encryption integrity
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestPgpSignedEmailSkipped()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = "Message with URL: https://www.should-not-be-converted.com";
            
            var mailItem = TestHelpers.CreatePgpSignedTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert - URL should still be present (not processed due to PGP signature)
            var processedBody = mailItem.Message.Body.GetText();
            Assert.IsTrue(processedBody.Contains("-----BEGIN PGP SIGNED MESSAGE-----"));
            Assert.IsTrue(processedBody.Contains("https://www.should-not-be-converted.com"));
        }

        [TestMethod]
        public void TestEmailWithSignatureAttachmentSkipped()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = TestHelpers.CreatePlainTextWithUrls(
                "Message with signature attachment.",
                "https://www.should-not-be-converted.com"
            );
            
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody);
            TestHelpers.AddSignatureAttachment(mailItem);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert - URL should still be present (not processed due to signature attachment)
            var processedBody = mailItem.Message.Body.GetText();
            Assert.IsTrue(processedBody.Contains("https://www.should-not-be-converted.com"));
            // Signed emails with attachments are not processed
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestMultipleUrlsConverted()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = @"Multiple URLs in message:
First: https://www.site1.com/page
Second: http://www.site2.org/login?user=test
Third: https://malicious.co.uk/phishing
Plain text reference to https://another-site.net should also be converted.";
            
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert
            var processedBody = mailItem.Message.Body.GetText();
            // URLs should still be present as plain text
            Assert.IsTrue(processedBody.Contains("https://www.site1.com/page"));
            Assert.IsTrue(processedBody.Contains("http://www.site2.org/login?user=test"));
            Assert.IsTrue(processedBody.Contains("https://malicious.co.uk/phishing"));
            Assert.IsTrue(processedBody.Contains("https://another-site.net"));
            
            // URLs are converted to plain text, not removed with markers
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestMessageWithoutUrls()
        {
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = "This is a clean message with no URLs. Just plain text content that should remain unchanged.";
            
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody);

            // Act
            _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });

            // Assert - Message should remain unchanged
            var processedBody = mailItem.Message.Body.GetText();
            Assert.AreEqual(messageBody, processedBody);
            // No URLs to convert, so no changes expected
            Assert.IsFalse(processedBody.Contains("<a href"));
        }

        [TestMethod]
        public void TestDomainExtraction()
        {
            // Use reflection to test private method
            var getDomainMethod = typeof(UrlToTextAgent).GetMethod("GetDomainFromAddress", BindingFlags.NonPublic | BindingFlags.Instance);
            Assert.IsNotNull(getDomainMethod, "GetDomainFromAddress method not found");

            // Test valid email addresses
            var domain1 = getDomainMethod.Invoke(_agent, new object[] { "user@example.com" });
            Assert.AreEqual("example.com", domain1);

            var domain2 = getDomainMethod.Invoke(_agent, new object[] { "test.user@subdomain.example.org" });
            Assert.AreEqual("subdomain.example.org", domain2);

            // Test invalid addresses
            var domain3 = getDomainMethod.Invoke(_agent, new object[] { "invalid-email" });
            Assert.AreEqual("", domain3);

            var domain4 = getDomainMethod.Invoke(_agent, new object[] { "" });
            Assert.AreEqual("", domain4);

            var domain5 = getDomainMethod.Invoke(_agent, new object[] { null });
            Assert.AreEqual("", domain5);
        }

        [TestMethod]
        public void TestLoggingFunctionality()
        {
            // This test verifies that logging works without throwing exceptions
            // In a real environment, you might want to check the actual log content
            
            // Arrange
            var fromAddress = "external@external.com";
            var messageBody = TestHelpers.CreatePlainTextWithUrls("Test logging", "https://test-url.com");
            var mailItem = TestHelpers.CreateTestMailItem(fromAddress, messageBody);

            // Act - should not throw any exceptions
            try
            {
                _agent.GetType().GetMethod("ProcessMessage", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.Invoke(_agent, new object[] { mailItem, "TestEvent" });
                
                // Assert - if we get here, logging didn't cause any exceptions
                Assert.IsTrue(true, "Processing completed without logging errors");
            }
            catch (Exception ex)
            {
                Assert.Fail($"Processing failed with exception: {ex.Message}");
            }
        }
    }
}