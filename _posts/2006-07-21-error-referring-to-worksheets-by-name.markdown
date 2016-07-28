---
layout: post
title:  "Error when referring to Excel sheets collection by sheet name"
date:   2006-07-21 10:15:00 +1000
tags:   [vba, excel, worksheet, collection, bug]
---

I've just had a frustrating couple of hours dealing with what appears to
be a bug in the Excel Sheets collection when accessed from an Access VBA
module. The application is an Access 2000 database reading data from an
Excel workbook. The method employed is to loop through the table
definition and retrieve the value of any Named Ranges from the work book
which match the name of each field in the table. This works fine.

But a new version of the workbook was issued to users with some of the
named range definitions missing. So to compensate for this, I added a
correction table to the database with three text fields - FieldName,
Sheet (e.g. "xyz") and Cell (e.g. "A5"). When the load routine fails to
find the required field in the workbook, the following code segment was
intended to retrieve the value from the workbook:

```vb
varData = wb.Sheets(rst!Sheet).Range(rst!Cell).Value
```

This failed with **Run-time error '13': Type mismatch**. Subsequent
testing in the immediate window showed the culprit to be accessing the
Sheets collection with wb.Sheets(18).Name returning "xyz" but
wb.Sheets("xyz").Name failing with Error 13.

I don't get this error from VBA running in the Excel client, only when
Excel is an automation client and the code is running within the Access
client. It also applies to the Worksheets collection as well (as you
would expect).

The workaround was to change the correction table to be two fields -
FieldName & RefersTo (which is Sheet and Cell combined e.g. "xyz!A5")
and to use the following which works without any problems:

```vb
varData = appExcel.Range(rst!RefersTo).Value
```
