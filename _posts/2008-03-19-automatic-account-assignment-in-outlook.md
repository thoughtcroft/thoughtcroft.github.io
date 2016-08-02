---
layout:   post
title:    "Automatic email account assignment in Outlook"
date:     2008-03-19
category: vba
tags:
  - email
  - mapi
  - outlook
  - redemption
  - vba
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

{% gist thoughtcroft/1497ecbab5d77abe43323303485b7d95 clsNewMailHandler.vb %}

Next create a new standard Module called `basMailRoutines` and import this
code:

{% gist thoughtcroft/1497ecbab5d77abe43323303485b7d95 basMailRoutines.vb %}

Finally, add the following code to the `ThisOutlookSession` object:

{% gist thoughtcroft/1497ecbab5d77abe43323303485b7d95 ThisOutlookSession.vb %}

Restart Outlook and you are in business! Obviously you could rearrange
this code to suit your own purpose and condense some of the code into
the one class module but I find this modularisation makes the code much
easier to understand and manage.
