---
layout:   post
title:    "How to tell if a file is a Word document"
date:     2010-04-01
category: vba
tags:
 - file
 - vba
 - word
---

I am building a tool that needs to do something to all the Word
documents in a set of folders. There are other sorts of files in these
folders so I need to filter the list. I didn't want to hard code the
file extensions that constitute Word documents because these change
according to the version and will do so in future. So I decided to
develop a function that is cross-version compatible.

To determine the documents that Microsoft Word can open and update
natively, I am using the `Filters` collection from the
`Application.FileDialog` object which is available in Word version 2002
onwards. This is the list of extensions that can be selected in the File
Open dialog to filter the files that can be opened in Word. Seems to
work like a charm. For earlier versions, you'll need to hardcode the
extensions.

{% gist thoughtcroft/722d4f97b0622726a98f8e0d3f341dd1 %}
