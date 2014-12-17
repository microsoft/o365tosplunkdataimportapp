using Microsoft.Office365.ReportingWebServiceClient;
using Splunk.ModularInputs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace o365ToSplunkDataImport
{
    public class Program : ModularInput
    {
        private const string ConstantReportName = "reportName";
        private const string ConstantEmailAddress = "email";
        private const string ConstantPassword = "password";
        private const string ConstantStartDate = "startDate";
        private const string ConstantEndDate = "endDate";

        #region Private Methods

        private async Task<string> GetConfigurationValue(InputDefinition definition, string keyName, EventWriter writer, bool log = true)
        {
            Parameter parameter;
            if (definition.Parameters.TryGetValue(keyName, out parameter))
            {
                await writer.LogAsync(Severity.Info, string.Format("Value for [{0}] retrieved successfully: {1}", keyName, log ? parameter.ToString() : "# Value not logged #"));
                return parameter.ToString();
            }
            else
                return null;
        }

        #endregion Private Methods

        public static int Main(string[] args)
        {
#if DEBUGNOATTACH
            return Run<Program>(args, DebuggerAttachPoints.None);
#elif DEBUG
            return Run<Program>(args, DebuggerAttachPoints.StreamEvents); //DebuggerAttachPoints.ValidateArguments | 
#else
            return Run<Program>(args);
#endif
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
                                Name = ConstantReportName,
                                Title = "Office 365 report name",
                                Description = "Note: Your Office 365 subscription might not include all of the reports listed here. To see which reports are part of your subscription, check the Reports section of Office 365",
                                RequiredOnCreate = true,
                                RequiredOnEdit = true,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = ConstantEmailAddress,
                                Title = "Office 365 email address",
                                Description = "The email address you use to sign in to Office 365 for business",
                                RequiredOnCreate = true,
                                RequiredOnEdit = true,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = ConstantPassword,
                                Title = "Office 365 password",
                                Description = string.Empty,
                                RequiredOnCreate = true,
                                RequiredOnEdit = true,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = ConstantStartDate,
                                Title = "From date",
                                Description = "(Optional) The earliest date you want to fetch data from. Date time format: yyyy-MM-dd hh:mm:ss, for example: 2014-09-01 08:00:00",
                                RequiredOnCreate = false,
                                RequiredOnEdit = false,
                                DataType = DataType.String
                            },
                            new Argument {
                                Name = ConstantEndDate,
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
                #region Get param values

                //Adding the validate value at the end to differentiate it from the real stream later created while streaming data
                string streamName = validation.Name + "validate";
                string reportName = validation.Parameters[ConstantReportName].ToString();
                string emailAddress = validation.Parameters[ConstantEmailAddress].ToString();
                string password = validation.Parameters[ConstantPassword].ToString();
                DateTime startDateTime = validation.Parameters[ConstantStartDate].ToString().TryParseDateTime(DateTime.MinValue);
                DateTime endDateTime = validation.Parameters[ConstantEndDate].ToString().TryParseDateTime(DateTime.MaxValue);

                #endregion Get param values

                if (startDateTime > endDateTime)
                {
                    errorMessage = "From date must be less than or equal to To date";
                    return false;
                }

                ReportingContext context = new ReportingContext();
                context.UserName = emailAddress;
                context.Password = password;
                context.FromDateTime = startDateTime;
                context.ToDateTime = endDateTime;
                //TODO: Need the EventWriter instance to log stuff here
                context.SetLogger(new DefaultLogger());

                ReportingStream stream = new ReportingStream(context, reportName, streamName);
                
                bool res = stream.ValidateAccessToReport();
                if(res)
                {
                    errorMessage = "";
                    stream.ClearProgress();
                }
                else
                {
                    errorMessage = string.Format("An error occured while validating your crendentials against report: {0}", reportName);
                }
                return res;
                    
            }
            catch (Exception e)
            {
                errorMessage = e.Message;
                return false;
            }
        }

        public override async Task StreamEventsAsync(InputDefinition inputDefinition, EventWriter eventWriter)
        {
            #region Get param values

            string streamName = inputDefinition.Name;
            await eventWriter.LogAsync(Severity.Info, string.Format("Name of Stanza is : {0}", inputDefinition.Name));

            string reportName = await GetConfigurationValue(inputDefinition, ConstantReportName, eventWriter);
            string emailAddress = await GetConfigurationValue(inputDefinition, ConstantEmailAddress, eventWriter);
            string password = await GetConfigurationValue(inputDefinition, ConstantPassword, eventWriter, false);
            string start = await GetConfigurationValue(inputDefinition, ConstantStartDate, eventWriter);
            DateTime startDateTime = start.TryParseDateTime(DateTime.MinValue);
            string end = await GetConfigurationValue(inputDefinition, ConstantEndDate, eventWriter);
            DateTime endDateTime = end.TryParseDateTime(DateTime.MinValue);

            #endregion Get param values

            ReportingContext context = new ReportingContext();
            context.UserName = emailAddress;
            context.Password = password;
            context.FromDateTime = startDateTime;
            context.ToDateTime = endDateTime;
            context.SetLogger(new SplunkTraceLogger(eventWriter));

            IReportVisitor visitor = new SplunkReportVisitor(streamName, eventWriter);

            while (true)
            {
                await Task.Delay(1000);

                ReportingStream stream = new ReportingStream(context, reportName, streamName);
                stream.RetrieveData(visitor);
            }
        }
    }

    internal static class Extensions
    {
        public static Guid TryParseGuid(this string value, Guid defaultValue)
        {
            Guid result;
            if (Guid.TryParse(value, out result))
            {
                return result;
            }

            return defaultValue;
        }

        public static DateTime TryParseDateTime(this string value, DateTime defaultValue)
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

        public static int TryParseInt(this string value, int defaultValue)
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
    }
}