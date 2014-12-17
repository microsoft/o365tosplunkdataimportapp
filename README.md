# Office 365 to Splunk data import app

The Office 365 data Splunk app enables data analysts and IT administrators to import the data they need to get their organization more productive and finally makes Office 365 data available to Splunk 

## Pre-requisites

* [Splunk 6.x Enterprise](http://www.splunk.com/)
* Windows 7 and later, and Windows Server 2008 and later

Note: Tested with Splunk Enterprise 6.0, 6.1 and 6.2

## Dependencies

* [Splunk Modular Input SDK for .Net](https://www.nuget.org/packages/Splunk.ModularInputs/)
* [Office 365 reporting web service client library](http://www.nuget.org/packages/Microsoft.Office365.ReportingWebServiceClient/)

Note: Office 365 reporting web service client library is also an open source project: <https://github.com/Microsoft/o365rwsclient>

## Known limitations

* ~~Currently, the 'StreamEvents' method will be called once when Splunk run the executable, which means that it does not yet continuously pull data from Office 365 reporting web service. Some extra logic should be implemented to hanlde this.~~ We have fixed this in 1.1 beta!
* We have not yet created all the report proxy classes, which means that you may not be able to fetch data from some of your tenant's subscription.
* We save a ".progress" file in the bin folder of the application in order to track the timestamp the last fetch occured. The current issue is that the progress is identified with the Splunk stream name, which means that if you created a stream with name, say "MailboxUsageDetail1", fetched data and then deleted that stream, the progress file is NOT currently deleted, so if you create a new stream with the same name, "MailboxUsageDetail1", the software will pick up the progress file from the previous file!!!
* If you create a data stream with an "End Date", the software will keep running forever even if it does not fetch data anymore. To save resources on your Splunk server, you should consider disabling a stream that has finished fetching data.

## Building and deploying at dev box

The project file has build events that pushes the projects file to a %SPLUNK_HOME% location, which you will need to create in your system environment variables

![PostBuildEvent](/doc/AddSplunkHomeVariable.png?raw=true)

## Feedback

* [StackOverflow](http://stackoverflow.com/questions/tagged/o365tosplunkapp)
* [Create a tagged StackOverflow question](http://stackoverflow.com/questions/ask?tags=o365tosplunkapp)
* [Yammer IT Pro Network group](http://aka.ms/o365tosplunkappfeedback)