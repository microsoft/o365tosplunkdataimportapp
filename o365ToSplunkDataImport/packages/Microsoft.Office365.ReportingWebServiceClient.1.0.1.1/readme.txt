Microsoft Office 365 Reporting web service client library
Copyright © 2014 Microsoft Corporation
Authors: Julien Pierre (Product Manager), Toney Sui (Engineer)

### Getting started

Assuming you are building a standard C# console application, paste the code below in the Main method:

ReportingContext context = new ReportingContext();
context.UserName = @"YOUR O365 EMAIL ADDRESS";
context.Password = @"YOUR O365 PASSWORD";
context.FromDateTime = DateTime.MinValue;
context.ToDateTime = DateTime.MinValue;
context.SetLogger(new DefaultLogger());

IReportVisitor visitor = new DefaultReportVisitor();

ReportingStream stream = new ReportingStream(context, "MailboxActivityDaily", "stream1");
stream.RetrieveData(visitor);

// Clear the progress, so that the next time you run it will get same result
stream.ClearProgress();

Console.WriteLine("Press Any Key...");
Console.ReadKey();


Enjoy!
Source code + information >> https://github.com/Microsoft/o365rwsclient