using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;

namespace Microsoft.Splunk.O365Reporting
{
    public class SplunkReportVisitor : IReportVisitor
    {
        private string streamName;

        public SplunkReportVisitor(string stream)
        {
            this.streamName = stream;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="report"></param>
        public override void VisitReport(ReportObject report)
        {
            using (EventStreamWriter writer = new EventStreamWriter())
            {
                writer.Write(new EventElement
                {
                    Time = report.Date,
                    Source = this.streamName,
                    Data = report.ConvertToXml()
                });
            }
        }

        /// <summary>
        ///
        /// </summary>
        public override void VisitBatchReport()
        {
            SystemLogger.Write(string.Format("VisitBatchReport: {0}", this.reportObjectList.Count));

            using (EventStreamWriter writer = new EventStreamWriter())
            {
                foreach (ReportObject report in reportObjectList)
                {
                    writer.Write(new EventElement
                    {
                        Time = report.Date,
                        Source = this.streamName,
                        Data = report.ConvertToXml()
                    });
                }
            }
        }
    }
}