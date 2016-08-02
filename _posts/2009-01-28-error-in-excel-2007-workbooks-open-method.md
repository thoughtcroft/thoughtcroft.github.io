---
layout:   post
title:    "Error in Excel 2007 workbooks open method"
date:     2009-01-28
category: vba
tags:
  - access
  - excel
  - password
  - workbook
  - vba
---

I have discovered a disturbing "feature" in Excel 2007 VBA in the way
that the Workbooks.Open method handles xlsx format files versus the
legacy xls format file.

If a file has a "workbook protection" password set, then the following
statement will succeed for an xls file but will fail for an xlsx file
with run time error "1004: The password you supplied is not correct."

```vb
set wkb = Application.Workbooks.Open("file.???", , , , "password")
```

Why supply a password in this call? I am running Excel as an unattended
automation server to load data captured in workbooks submitted from a
website into an Access database. Sometimes the users apply a "file open"
password which will cause Excel to display a dialog box and wait for
someone to supply the password if no password is supplied as an argument
to the Open method - very bad! On the other hand, if the password
argument is present and is incorrect then the above 1004 error is
generated and I can then deal with it in code. If there isn't a "file
open" password on the file, the Open method ignores the password
argument.

However, all my files have a "workbook protection" password set and some
users submit the files in the new xlsx format. This causes me to discard
the file when it should be able to be opened. By the way, even if the
"workbook protection" password is "password" it generates this error.
And if there is a "workbook protection" password on an xlsx file and no
password argument is provided to the Open method, it opens the file
without any trouble.

So looks like Microsoft have messed up the Workbooks Open method in the
case where:

* The file has a "workbook protection" password
* The file is in the new xlsx format
* A password argument is supplied in the call to the Workbooks.Open method

This has been reported to Microsoft via Excel MVP [Ron de
Bruin](http://www.rondebruin.nl/tips.htm).
