# Office 365 to Splunk data import app - o365ToSplunkDataImportApp

The Office 365 data Splunk app enables data analysts and IT administrators to import the data they need to get their organization more productive and finally makes Office 365 data available to third party BI platforms 

## Pre-requisites

* [Splunk 6.x Enterprise](http://www.splunk.com/)

Note: Currently tested with Splunk Enterprise 6.0

## Dependencies

* [Splunk SDK for .Net](http://www.nuget.org/packages/SplunkSDK/)
* [Office 365 reporting web service client library](http://www.nuget.org/packages/Microsoft.Office365.ReportingWebServiceClient/)

Note: Office 365 reporting web service client library is also an open source project: [https://github.com/Microsoft/o365rwsclient](https://github.com/Microsoft/o365rwsclient)

## Building and deploying

For simplicity and debugging reasons we decided to leverage Visual Studio post-build event to copy the required files directly to the Splunk server installation path, which is currently assumed to be "C:\Program Files\Splunk\".

![PostBuildEvent](/doc/PostBuildEventCommandLine.png?raw=true)

So we assume you are building the project on the machine that has Splunk Enterprise installed. If not you should modify the post-build event.

# Feedback

* [StackOverflow](http://stackoverflow.com/questions/tagged/o365tosplunkapp)
* [Create a tagged StackOverflow question](http://stackoverflow.com/questions/ask?tags=o365tosplunkapp)
* [Yammer IT Pro Network group](http://aka.ms/o365tosplunkappfeedback)