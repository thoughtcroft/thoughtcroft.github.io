---
layout: post
title:  "Automatic email account assignment in Outlook"
date:   2008-03-19 15:33:00 +1000
tags:   [vba, email, outlook, account]
---

I use Google Apps to host my family's and my private company's email but
I manage my email through Outlook (because I prefer to operate
off-line and I need to synchronise with my Nokia N95 phone).

Google Mail is fantastic but the issue I had to grapple with relates to
how email is pulled into Outlook. Email sent to warren@[personal] and
warren@[business] ends up arriving in the one inbox and is automatically
assigned to the default email account in Outlook ([business] in this
case).

When I reply to an email I want it to be sent through the correct
account so that the sender and reply addresses match the right context.
I can do that manually using the Accounts button followed by the Send
button but there is another (better) way using VBA.

First download and install the excellent
[Redemption](http://www.dimastr.com/redemption/home.htm) COM Library written
by Dmitry Streblechenko. This will expose the required properties of the
Outlook object model without triggering the user confirmation dialog
introduced by the Outlook Security Patch. It also provides access to
many useful MAPI properties not available through the standard Outlook
object model.

The following code is triggered whenever new mail arrives in Outlook. In
essence, I look for certain phrases that indicate that the mail has been
sent to my [personal] address and then change the mail account for that
message to match. Subsequently, when I reply to that message, I don't
need to choose which address to send it from as that has already been
selected.

It works 95% of the time with misses probably due to timing issues and
possible conflicts with Outlook rules. The NewMailEx event is perhaps
not guaranteed to fire (despite what the documentation says) and so
sometimes the account is left set to the incorrect one but I am happy
enough with the result. The techniques employed here could be used for
other new mail triggered actions.

First, create a new Class module in your Outlook VBA project called
`clsNewMailHandler`

```vb
Option Explicit

Public WithEvents oApp As Outlook.Application
Const TC_BAINSWORLD_ACCOUNT = "bainsworld"

Private Sub Class_Terminate()
   Set oApp = Nothing
End Sub

Private Sub oApp_NewMailEx(ByVal EntryIDCollection As String)

' This will be called whenever we receive new mail so
' process each item to determine if we should alter
' the account - do we need to worry about conflicts with Rules?

   Dim astrEntryIDs() As String
   Dim objItem As Object
   Dim varEntryID As Variant

   astrEntryIDs = Split(EntryIDCollection, ",")
   For Each varEntryID In astrEntryIDs
       Set objItem = oApp.Session.GetItemFromID(varEntryID)
       If objItem.Class = olMail Then
           ' Only call this for MailItems - can be ReadReceipts
           ' too which are class olReport
           Call SetEmailAccount(objItem)
       End If
   Next varEntryID
   Set objItem = Nothing
End Sub

Private Sub SetEmailAccount(ByRef oItem As MailItem)

' This code will check if the item is of interest to
' us and if so will update the account property accordingly

' Check if this was sent to a 'bainsworld' address
   If CheckMessageRecipient(oItem, TC_BAINSWORLD_ACCOUNT, False) Then
       ' Yes it was - change the account
       Call SetMessageAccount(oItem, TC_BAINSWORLD_ACCOUNT, True)
   End If
End Sub

Private Sub Class_Initialize()
   Set oApp = Application
End Sub
```

Next create a new standard Module called `basMailRoutines` and import this
code:

```vb
Option Explicit

Private Const PR_HEADERS = &H7D001E
Private Const PR_ACCOUNT = &H80F8001E

Public Function CheckMessageRecipient( _
      ByRef oItem As MailItem, _
      ByVal strMatch As String, _
      Optional ByVal blnExact As Boolean = False) As Boolean

' Check if the supplied string matches the recipient
' of the email.  We use the internet headers and check
' the first part of the string if we can.  The match
' can be made exact or not

   Const TC_HEADER_START As String = "Delivered-To:"
   Const TC_HEADER_END As String = "Received:"

   Dim strHeader As String
   Dim intStart As Integer
   Dim intEnd As Integer
   Dim strRecipient As String

   ' First get the header and see if it makes sense
   strHeader = GetInternetHeaders(oItem)
   intStart = InStr(1, strHeader, TC_HEADER_START, vbTextCompare)
   If intStart = 0 Then intStart = 1
   intEnd = InStr(intStart, strHeader, vbCrLf & TC_HEADER_END, vbTextCompare)

   If intEnd = 0 Then
       ' The headers are unreliable so just check the whole string
       strRecipient = strHeader
   Else
       ' Found headers so grab the recipient data
       strRecipient = Trim$(Mid$(strHeader, intStart + Len(TC_HEADER_START), _
                                 intEnd - (intStart + Len(TC_HEADER_START))))
   End If

   ' Now undertake the check
   If blnExact Then
       CheckMessageRecipient = (strRecipient = strMatch)
   Else
       CheckMessageRecipient = (InStr(1, strRecipient, strMatch, vbTextCompare) > 0)
   End If
End Function

Public Sub SetMessageAccount(ByRef oItem As MailItem, _
                            ByVal strAccount As String, _
                            Optional blnSave As Boolean = True)

    Dim rMailItem                  As Redemption.RDOMail
    Dim rSession                   As Redemption.RDOSession
    Dim rAccount                   As Redemption.RDOAccount

    ' Use a RDO Session object to locate the account
    ' that we are interested in

    Set rSession = New Redemption.RDOSession
    rSession.MAPIOBJECT = Application.Session.MAPIOBJECT
    Set rAccount = rSession.Accounts(strAccount)
  
    ' Now use the RDO Mail object to change the account
    ' to the one we require

    Set rMailItem = rSession.GetMessageFromID(oItem.EntryID)
    rMailItem.Account = rAccount
    If blnSave Then
        ' They want us to force a save to the mail object
        rMailItem.Subject = rMailItem.Subject
        rMailItem.Save
    End If
    Set rMailItem = Nothing
    Set rAccount = Nothing
    Set rSession = Nothing
End Sub

Public Function GetInternetHeaders(ByRef oItem As MailItem) As String

   Dim rUtils As Redemption.MAPIUtils

   ' Return the internet header of a message
   Set rUtils = New Redemption.MAPIUtils
   GetInternetHeaders = rUtils.HrGetOneProp(oItem.MAPIOBJECT, PR_HEADERS)
   Set rUtils = Nothing
End Function
```

Finally, add the following code to the `ThisOutlookSession` object:

```vb
Dim MyNewMailHandler As clsNewMailHandler

Private Sub Application_Quit()
   Set MyNewMailHandler = Nothing
End Sub

Private Sub Application_Startup()
   Set MyNewMailHandler = New clsNewMailHandler
End Sub
```

Restart Outlook and you are in business! Obviously you could rearrange
this code to suit your own purpose and condense some of the code into
the one class module but I find this modularisation makes the code much
easier to understand and manage.
