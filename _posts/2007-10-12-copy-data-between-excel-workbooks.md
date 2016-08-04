---
layout:   post
title:    "Generic function to copy Excel data between workbooks"
date:     2007-10-12
category: vba
tags:
  - excel
  - password
  - upload
  - vba
  - workbook
---

I needed a way to ensure that some Excel workbooks that were being
completed by customers and uploaded to a server for analysis and
subsequent loading into a database, were not being tampered with.
Certainly in Excel version up to and including 2003 it is not possible
to prevent a determined user from "cracking" the Worksheet and Workbook
protection passwords and hence to modifying any formulae in the
Workbook.

My solution was to create a new, clean copy of the workbook from a
master template and to then copy the customer entered data from the
submitted book to the new clean book. The following CopyExcelData sub
does the bulk of the work - the comments explain what is going on.

It does take some time to do this as it operates cell-by-cell (around 1
minute to copy 500 cells spread across 15 worksheets) but its not a bad
trade-off to be safe in the knowledge that the final result is
tamper-proof.

As a side benefit, any corrections needed to the customer workbook after
it has been issued can be made to the template version and files
received after that point will be automatically "upgraded".

{% gist thoughtcroft/fae2f2ce729f1dbe0f833a7855a82838 %}
