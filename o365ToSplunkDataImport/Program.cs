using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;

namespace Microsoft.Splunk.O365Reporting
{
    public class Program : ModularInput
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
                    Title = "Office 365 data import",
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

 
        public override bool Validate(Validation validation, out string errorMessage)
        {
            try
            {
                #region Get stanza values

                string streamName = validation.Name;

                string reportName = validation.Parameters[ConstantReportName].ToString();
                string emailAddress = validation.Parameters[ConstantEmailAddress].ToString();
                string password = validation.Parameters[ConstantReportName].ToString();
                string startDateTemp = validation.Parameters[ConstantStartDate].ToString();
                DateTime startDate = TryParseDateTime(startDateTemp, DateTime.MinValue);
                string endDateTemp = validation.Parameters[ConstantEndDate].ToString();
                DateTime endDate = TryParseDateTime(endDateTemp, DateTime.MaxValue);

                #endregion Get stanza values
            
                if (startDate > endDate)
                {
                    errorMessage = "startDate must be less than or equal to endDate";
                    return false;
                }

                ReportingContext context = new ReportingContext("https://reports.office365.com/ecp/reportingwebservice/reporting.svc");
                context.UserName = emailAddress;
                context.Password = password;
                context.FromDateTime = startDate;
                context.ToDateTime = endDate;

                ReportingStream stream = new ReportingStream(context, reportName, streamName);
                stream.ValidateAccessToReport();
            }
            catch(Exception e)
            {
                errorMessage = e.Message;
                return false;
            }

            errorMessage = "";
            return true;
        }


        public override async Task StreamEventsAsync(InputDefinition inputDefinition, EventWriter eventWriter)
        {
            #region Get stanza values

            await eventWriter.LogAsync(Severity.Info, string.Format("Name of Stanza is : {0}", inputDefinition.Name));

            string reportName = await GetConfigurationValue(inputDefinition, ConstantReportName, eventWriter);
            string emailAddress = await GetConfigurationValue(inputDefinition, ConstantEmailAddress, eventWriter);
            string password = await GetConfigurationValue(inputDefinition, ConstantPassword, eventWriter);

            await eventWriter.LogAsync(Severity.Info, await GetConfigurationValue(inputDefinition, ConstantStartDate, eventWriter));

            DateTime startDate = TryParseDateTime(await GetConfigurationValue(inputDefinition, ConstantStartDate, eventWriter), DateTime.MinValue);
            DateTime endDate = TryParseDateTime(await GetConfigurationValue(inputDefinition, ConstantEndDate, eventWriter), DateTime.MinValue);

            #endregion Get stanza values

            string streamName = inputDefinition.Name;

            ReportingContext context = new ReportingContext("https://reports.office365.com/ecp/reportingwebservice/reporting.svc");
            context.UserName = emailAddress;
            context.Password = password;
            context.FromDateTime = startDate;
            context.ToDateTime = endDate;
            context.SetLogger(new SplunkTraceLogger(eventWriter));

            IReportVisitor visitor = new SplunkReportVisitor(streamName, eventWriter);

            ReportingStream stream = new ReportingStream(context, reportName, streamName);
            stream.RetrieveData(visitor);
        }

        #region Private Methods

        private async Task<string> GetConfigurationValue(InputDefinition definition, string keyName, EventWriter writer)
        {
            Parameter parameter;
            if (definition.Parameters.TryGetValue("keyName", out parameter))
            {
                await writer.LogAsync(Severity.Info, string.Format("Value for [{0}] retrieved successfully.", keyName));
                return parameter.ToString();
            }
            throw new ArgumentException(string.Format("Value for [{0}] retrieve failed. Return empty string.", keyName));
        }

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

        #endregion Private Methods
    }
}
