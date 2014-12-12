using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;

namespace o365ToSplunkDataImport
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
            writer.LogAsync(Severity.Error, message).Wait();
        }

        public void LogDebug(string message)
        {
            writer.LogAsync(Severity.Debug, message).Wait();
        }

        public void LogFatal(string message)
        {
            writer.LogAsync(Severity.Fatal, message).Wait();
        }

        public void LogWarning(string message)
        {
            writer.LogAsync(Severity.Warning, message).Wait();
        }
    }
}