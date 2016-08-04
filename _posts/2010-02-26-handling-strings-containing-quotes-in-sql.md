---
layout:   post
title:    "Handling strings containing quotes in SQL"
date:     2010-02-26
category: vba
tags:
  - access
  - quote
  - sql
  - string
  - vba
---

Everyone is aware that quotes in strings used in SQL cause all sorts of
problems. As developers, it is our duty to ensure that an end user
never sees an SQL error because we didn't handle his input
appropriately!

But how the heck do you ensure that quotes are handled
correctly? The following function will take a string and return it
enclosed with quotes and all instances of quotes in the string will be
doubled up.

For example if I call QuotedString("This is a "string" example") then it
will return:

"This is ""a string"" example"

```vb
Public Function QuotedString( _
       strText As String) As String

    Const conQuoteChar = """"

    QuotedString = conQuoteChar _
                 & Replace$(strText, conQuoteChar, conQuoteChar & conQuoteChar) _
                 & conQuoteChar
End Function
```

NOTE: you should always check for specific methods designed to escape
strings in whatever framework you are using. I had to roll my own
because Microsoft Access.
