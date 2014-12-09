using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;

namespace Microsoft.Splunk.O365Reporting
{
    public class SplunkReportVisitor : IReportVisitor
    {
        private readonly string streamName;
        private readonly EventWriter writer;

        public SplunkReportVisitor(string stream, EventWriter writer)
        {
            this.streamName = stream;
            this.writer = writer;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="report"></param>
        public override async void VisitReport(ReportObject report)
        {
            await writer.QueueEventForWriting(new Event{Time = report.Date, Source=this.streamName,Data=report.ConvertToXml()});
        }

        /// <summary>
        ///
        /// </summary>
        public override void VisitBatchReport()
        {
            writer.LogAsync(Severity.Info, string.Format("VisitBatchReport: {0}", this.reportObjectList.Count)).Wait();

            foreach (ReportObject report in reportObjectList)
            {
                writer.QueueEventForWriting(new Event{Time=report.Date, Source=this.streamName, Data=report.ConvertToXml()}).Wait();
            }
        }
    }
}