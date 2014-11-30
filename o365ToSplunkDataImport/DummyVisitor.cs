using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office365.ReportingWebServiceClient;

namespace Microsoft.Splunk.O365Reporting
{
    public class DummyVisitor : IReportVisitor
    {
        public override void VisitReport(ReportObject report)
        {
        }

        public override void VisitBatchReport()
        {
        }
    }
}
