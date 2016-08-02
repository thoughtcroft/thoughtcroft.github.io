---
layout: post
title:  "Automation Manager Class"
date:   2008-06-19
tags:   [access, automation, excel, vba]
---

I needed a way to manage calls to different office automation servers in
a consistent fashion. These were mostly from Access to extract data from
a large number of Excel workbooks. Specifically what I wanted was a way
of managing:

* reuse of any existing instance of Excel or start an instance if there
  wasn't one

* save the state of the application - things like the calculation mode
  etc and restore them when finished

* work out whether to close the instance when finished (if we started
  it) or leave it (if we didn't)

* handle the strange automation errors that can occur and ensure that
  the instance is properly terminated in the case where an unrecoverable
  error has occurred

Further, I wanted to be able to use this for multiple automation
clients. The following class modules have served me well for this
purpose and I offer them here for those that may have a similar
requirement. Some of this is not for the faint hearted, so send me an
email if you need further explanation.

Typical calling method is as follows:

{% gist thoughtcroft/62e864dc99edb15023e03cd7a1654fd9 Example.vb %}

Create a new Class Module called AppState and copy the following code into it. This describes all the properties associated with an instance of an application that has been started to provide automation services.

{% gist thoughtcroft/62e864dc99edb15023e03cd7a1654fd9 AppState.vb %}

Create a Class Module called AppStateMgr and copy the following code into it. This provides the functions for managing instances of automation clients.

{% gist thoughtcroft/62e864dc99edb15023e03cd7a1654fd9 AppStateMgr.vb %}
