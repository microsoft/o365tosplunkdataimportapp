# Office 365 to Splunk data import app - o365ToSplunkDataImportApp

The Office 365 data Splunk app enables data analysts and IT administrators to import the data they need to get their organization more productive and finally makes Office 365 data available to third party BI platforms 

## Pre-requisites

* [Splunk 6.x Enterprise](http://www.splunk.com/)
* Windows 7 and later, and Windows Server 2008 and later

Note: Tested with Splunk Enterprise 6.0, 6.1 and 6.2

## Dependencies

* [Splunk SDK for .Net](http://www.nuget.org/packages/SplunkSDK/)
* [Office 365 reporting web service client library](http://www.nuget.org/packages/Microsoft.Office365.ReportingWebServiceClient/)

Note: Office 365 reporting web service client library is also an open source project: [https://github.com/Microsoft/o365rwsclient](https://github.com/Microsoft/o365rwsclient)

## Known limitations

* Currently, the 'StreamEvents' method will be called once when Splunk run the executable, which means that it does not yet continuously pull data from Office 365 reporting web service. Some extra logic should be implemented to hanlde this.
* We have not yet created all the report proxy classes, which means that you may not be able to fetch data from some of your tenant's subscription.

## Building and deploying at dev box

For simplicity and debugging reasons we decided to leverage Visual Studio post-build event to copy the required files directly to the Splunk server installation path, which is currently assumed to be "C:\Program Files\Splunk\".

![PostBuildEvent](/doc/PostBuildEventCommandLine.png?raw=true)

So we assume you are building the project on the machine that has Splunk Enterprise installed. If not you should modify the post-build event.

## Feedback

* [StackOverflow](http://stackoverflow.com/questions/tagged/o365tosplunkapp)
* [Create a tagged StackOverflow question](http://stackoverflow.com/questions/ask?tags=o365tosplunkapp)
* [Yammer IT Pro Network group](http://aka.ms/o365tosplunkappfeedback)