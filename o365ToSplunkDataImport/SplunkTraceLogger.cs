using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;

namespace Microsoft.Splunk.O365Reporting
{
    public class SplunkTraceLogger : ITraceLogger
    {
        public void LogInformation(string message)
        {
            SystemLogger.Write(message);
        }

        public void LogError(string message)
        {
            SystemLogger.Write(message);
        }
    }
}