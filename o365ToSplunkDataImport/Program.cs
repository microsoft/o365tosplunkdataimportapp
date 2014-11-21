using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;
using System;
using System.Collections.Generic;

namespace Microsoft.Splunk.O365Reporting
{
    internal class Program : Script
    {
        private const string ConstantReportName = "reportName";
        private const string ConstantEmailAddress = "email";
        private const string ConstantPassword = "password";
        private const string ConstantStartDate = "startDate";
        private const string ConstantEndDate = "endDate";

        public static int Main(string[] args)
        {
            return Run<Program>(args);
        }

        public override Scheme Scheme
        {
            get
            {
                return new Scheme
                {
                    Title = "Office 365 to Splunk data import",
                    Description = "Get reporting data from your Office 365 for business subscription",
                    StreamingMode = StreamingMode.Xml,
                    Endpoint =
                    {
                        Arguments = new List<Argument> {
                            new Argument {
                                Name = "reportName",
                                Title = "Office 365 report name",
                                Description = "Note: Your Office 365 subscription might not include all of the reports listed here. To see which reports are part of your subscription, check the Reports section of Office 365",
                                RequiredOnCreate = true,
                                RequiredOnEdit = true,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = "email",
                                Title = "Office 365 email address",
                                Description = "The email address you use to sign in to Office 365 for business",
                                RequiredOnCreate = true,
                                RequiredOnEdit = true,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = "password",
                                Title = "Office 365 password",
                                Description = string.Empty,
                                RequiredOnCreate = true,
                                RequiredOnEdit = true,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = "startDate",
                                Title = "From date",
                                Description = "(Optional) The earliest date you want to fetch data from. Date time format: yyyy-MM-dd hh:mm:ss, for example: 2014-09-01 08:00:00",
                                RequiredOnCreate = false,
                                RequiredOnEdit = false,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = "endDate",
                                Title = "To date",
                                Description = "(Optional) The latest date you want to fetch data from. Date time format: yyyy-MM-dd hh:mm:ss, for example: 2014-09-01 08:00:00",
                                RequiredOnCreate = false,
                                RequiredOnEdit = false,
                                DataType = DataType.String
                            }
                        }
                    }
                };
            }
        }

        public override void StreamEvents(InputDefinition inputDefinition)
        {
            #region Get stanza values

            Stanza stanza = inputDefinition.Stanza;
            SystemLogger.Write(string.Format("Name of Stanza is : {0}", stanza.Name));

            string reportName = GetConfigurationValue(stanza, ConstantReportName);
            string emailAddress = GetConfigurationValue(stanza, ConstantEmailAddress);
            string password = GetConfigurationValue(stanza, ConstantPassword);

            SystemLogger.Write(GetConfigurationValue(stanza, ConstantStartDate));

            DateTime startDate = TryParseDateTime(GetConfigurationValue(stanza, ConstantStartDate), DateTime.MinValue);
            DateTime endDate = TryParseDateTime(GetConfigurationValue(stanza, ConstantEndDate), DateTime.MinValue);

            #endregion Get stanza values

            string streamName = stanza.Name;

            ReportingContext context = new ReportingContext("https://reports.office365.com/ecp/reportingwebservice/reporting.svc");
            context.UserName = GetConfigurationValue(stanza, ConstantEmailAddress);
            context.Password = GetConfigurationValue(stanza, ConstantPassword);
            context.FromDateTime = TryParseDateTime(GetConfigurationValue(stanza, ConstantStartDate), DateTime.MinValue);
            context.ToDateTime = TryParseDateTime(GetConfigurationValue(stanza, ConstantEndDate), DateTime.MinValue);
            context.SetLogger(new SplunkTraceLogger());

            IReportVisitor visitor = new SplunkReportVisitor(streamName);

            ReportingStream stream = new ReportingStream(context, reportName, streamName);
            stream.RetrieveData(visitor);
        }

        #region Private Methods

        private string GetConfigurationValue(Stanza stanza, string keyName)
        {
            string value;
            if (stanza.SingleValueParameters.TryGetValue(
                keyName,
                out value))
            {
                SystemLogger.Write(string.Format("Value for [{0}] retrieved successfully.", keyName));
                return value;
            }

            SystemLogger.Write(string.Format("Value for [{0}] retrieved failed. Return empty string.", keyName));
            return string.Empty;
        }

        private static Guid TryParseGuid(string value, Guid defaultValue)
        {
            Guid result;
            if (Guid.TryParse(value, out result))
            {
                return result;
            }

            return defaultValue;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private static DateTime TryParseDateTime(string value, DateTime defaultValue)
        {
            DateTime result;
            if (DateTime.TryParse(value, out result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private static int TryParseInt(string value, int defaultValue)
        {
            int result;
            if (int.TryParse(value, out result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
        }

        #endregion Private Methods
    }
}