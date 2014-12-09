using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;

namespace Microsoft.Splunk.O365Reporting
{
    public class SplunkTraceLogger : ITraceLogger
    {
        private EventWriter writer;

        public SplunkTraceLogger(EventWriter eventWriter)
        {
            this.writer = eventWriter;
        }
        public void LogInformation(string message)
        {
            writer.LogAsync(Severity.Info, message).Wait();
        }

        public void LogError(string message)
        {
            writer.LogAsync(Severity.Info, message).Wait();
        }
    }
}