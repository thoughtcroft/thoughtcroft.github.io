---
layout:   post
title:    "How to tell if an Access Report was printed"
date:     2007-03-01
category: vba
tags:
  - access
  - print
  - report
  - vba
---

I needed to be able to determine if an Access Report had been actually
sent to the printer (as opposed to just previewed on screen) so that I
could update a log recording the fact. This is useful for tracking
whether or not a letter has been sent to a customer without requiring
the user to click on a separate "log" button.

After some research on the 'net I discovered that a lot of the published
solutions only half-solved the problem. Microsoft themselves got it
completely wrong in this [knowledgebase
article](http://support.microsoft.com/kb/q154894/)!

The trick is understanding the different events that are fired when an
Access report is opened. Supposing the ReportHeader section is visible,
its Print event will be fired when the report is generated in Preview
mode and again every time the report is sent to the printer. To guard
against the situation where the report is sent direct to the printer
without first being opened in preview, we also need to look at the
Activate event which will be fired when the report is opened in preview
mode. Because the Activate event also fires whenever we switch back the
preview from another window, we also need to track the Deactivate event
to know that we have switched away from the preview.

And so that I don't have to add code to every report that I want to
track for all these events, I will define a class that sinks the events
in the report's open events as follows:

{% gist thoughtcroft/1b8a90d35411bb65ccc452765e0a01da ReportPrintStatus.vb %}

Now it becomes a simple matter to have a report work out if it was
printed or just previewed by inserting the following lines in the
report's code module. Note that this was in Access 97 - in later
versions I could raise a ReportPrinted event.

{% gist thoughtcroft/1b8a90d35411bb65ccc452765e0a01da ExampleUse.vb %}
