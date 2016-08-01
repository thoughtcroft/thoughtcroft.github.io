---
layout: post
title:  "How to tell if an Access Report was printed"
date:   2007-03-01
tags:   [access, report, vba]
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

```vb
' ReportPrintStatus class definition

' Use the hidden object type for a section so we can sink the events
' (thanks to Stephen Lebans for this tip - see www.lebans.com)

Private WithEvents mrpt As Access.Report
Private WithEvents msecReportHeader As Access.[_SectionInReport]
Private mintCounter As Integer

'--------------------------------
' Public Properties and Methods
'--------------------------------

Public Property Set Report(rpt As Access.Report)
   ' Sink the event handling for this report
   Const strEventKey As String = "[Event Procedure]"
   Set mrpt = rpt
   With mrpt
      ' If we don't populate these properties, the events will
      ' never fire in the report and we will be sunk!
      .OnActivate = strEventKey
      .OnClose = strEventKey
      .OnDeactivate = strEventKey
      ' Note, we assume this section exists - if not, it won't work
      Set msecReportHeader = .Section(acHeader)
      msecReportHeader.OnPrint = strEventKey
   End With
End Property

Public Property Get Printed() As Boolean
   ' Did they print this report?
   Printed = (mintCounter >= 1)
End Property

Public Sub Term()
   ' If we don't destroy these objects here, we risk an Access GPF!
   On Error Resume Next
   Set msecReportHeader = Nothing
   Set mrpt = Nothing
End Sub

'--------------------------------
' Event Procedures
'--------------------------------

Private Sub mrpt_Activate()
   ' This occurs if we open the report in print preview and also when
   ' we switch back to the previewed report in which case incremented
   ' by deactivate event
   mintCounter = mintCounter - 1
End Sub

Private Sub mrpt_Close()
   ' This occurs when the report closes so ensure we destroy objects
   ' to prevent an Access GPF
   Me.Term
End Sub

Private Sub mrpt_Deactivate()
   ' Called when we close report from preview or if we switch out of
   ' preview mode to another window in which case decremented by
   ' associated activate event
   mintCounter = mintCounter + 1
End Sub

Private Sub msecReportHeader_Print(Cancel As Integer, PrintCount As
Integer)
   ' Increment our counter occurs once for every time we print and also
   ' the first time we open in preview mode
   mintCounter = mintCounter + 1
End Sub
```

Now it becomes a simple matter to have a report work out if it was
printed or just previewed by inserting the following lines in the
report's code module. Note that this was in Access 97 - in later
versions I could raise a ReportPrinted event.

```vb
Private rps As ReportPrintStatus

Private Sub Report_Close()
   If rps.Printed Then
      ' Do something
    End If
End Sub

Private Sub Report_Open(Cancel As Integer)
   ' Sink the reports events so we can determine if it was printed or
not
   Set rps = New ReportPrintStatus
   Set rps.Report = Me
End Sub
```
