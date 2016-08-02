---
layout:   post
title:    "Strip different types of characters from a string"
date:     2008-06-18
category: vba
tags:
 - excel
 - string
 - vba
---

I realised that one of my earlier posts Generic Function to Copy Excel
Data references a function called 'tcStripChars' in order to remove
control characters from a cell value. This was required to prevent Excel
97 crashing when copying the .Value2 property of a cell into another
cell whenever there were control characters in the value e.g. LFCR.

The function is useful for other situations as well so here it is.

Note the use of the VBA6 conditional compilation constant to
generate appropriate code for your version of VBA using Enumerated Types
(VBA6) or Public Constants (pre VBA6).

{% gist thoughtcroft/12a4d5333d0a5f80f1fef56ba6ad7d6c %}
