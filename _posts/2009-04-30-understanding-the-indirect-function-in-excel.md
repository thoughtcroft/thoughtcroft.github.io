---
layout: post
title:  "Understanding the INDIRECT function in Excel"
date:   2009-04-30
tags:   [excel]
---

This is not a VBA issue but a recent discussion on LinkedIn brought up
the question of how to use INDIRECT to construct a reference to a named
range in an external sheet. I have been using this very useful function
for years to programmatically construct references based on different
parameters which resolve to a specific workbook, sheet, name etc. Here
is some information on how to properly construct a text based reference.

INDIRECT returns a reference specified by a text string and the syntax
is `INDIRECT(ref_text [, a1])`.

*ref_text* is either

* a reference to a cell containing an A1-style reference
* an R1C1-style reference
* a name defined as a reference; or
* a reference to a cell as a text string.

An error in the reference returns #REF!.

*a1* specifies the type of reference

* TRUE or omitted means A1-style
* FALSE means R1C1 style.

The INDIRECT function will return #REF! if the target workbook is not
open at the time you open the source workbook (closing the target after
it has been open doesn't cause this error as Excel remembers the value
until you force a recalc).

The proper construction of a text string based reference is as follows:

```
'path[file]sheet'!range
```

where the literal characters apostrophe ', square brackets [] and
exclamation mark ! delimit the following

* path = drive and folder where the file exists. If you leave this out
then it will use the workbook named file you have open no matter its
location on disk. If the path has a folder with spaces in it then you
must use the apostrophes.

* file = the name of the file with extension. Use of the apostrophes means
that this name can have spaces in it. If you leave out the apostrophes
then no spaces are allowed!

* sheet = the name of the worksheet. Only use the "[]" delimiters if a
sheet name is provided. No sheet name, no brackets around file!

If you don't provide a sheet name then the reference will default to the
first sheet in the workbook. For a named range, no sheet will work if
the named range has a scope of workbook or a scope of the first sheet.
Obviously, if you provide a sheet for a named range then it must have a
scope of that sheet or the reference doesn't exist and returns #REF!

I hope this is helpful to folks that are confused about this very useful
function.
