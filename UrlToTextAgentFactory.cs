using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;

namespace UrlToTextTransportAgent
{
    /// <summary>
    /// Factory class for creating UrlToTextAgent instances
    /// </summary>
    public class UrlToTextAgentFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new UrlToTextAgent();
        }
    }
}