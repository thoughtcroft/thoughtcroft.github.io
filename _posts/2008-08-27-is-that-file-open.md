---
layout: post
title:  "Is that file open?"
date:   2008-08-27
tags:   [file, windows, vba]
---

If you ever need to determine if a file is already open before you
attempt to do something with it in VBA, you will discover a blistering
array of code segments on the interweb which all seem to share the
characteristic that: **they fail to work as advertised**.

So time for me to add my version and I can state that it absolutely
works except where the opening application fails to lock the file.
AFAIK, that comment applies to text files opened in Notepad and any file
opened with an Office application that has its read-only attribute set.
However, since the main reason for testing if a file is already open is
so that you can do stuff to it without it causing a "file is already in
use" type of error, these two conditions don't matter since they don't
stop you doing anything with the file. If you don't believe me, try
this:

* Create an Excel file "test.xls"
* In Windows Explorer right click on the file and mark it as "read-only"
* Open the file in Excel
* Now go back to Windows Explorer and delete the file

Voila! No problem since it was read-only, Excel opened a temporary copy
of the file and so you do stuff to the original without it causing any
problems.

Now, back to the problem: how do you tell if a given file is in use?
Answer: use the VBA "open" statement to exclusively open the file and
trap any errors that may occur. We have to go one further though as the
"open" statement will not see any hidden files and will act as though
they don't exist. To overcome this we make some attribute changes and
trap any errors that may occur in relation to them as well.

{% gist thoughtcroft/9c750d95dad451bf61242f43d7ad2481 %}
