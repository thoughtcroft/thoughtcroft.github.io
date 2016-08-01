---
layout: post
title:  "Calculating compound annual growth rate (CAGR)"
date:   2005-11-30
tags:   [cagr, excel, udf, vba]
---

Microsoft Excel comes with a lot of inbuilt functions that can be used
in cell formulae and there are also a number of add-ins that provide
specialised sets of functions to support statistical, numerical and
financial analysis e.g. The Analysis ToolPak.

As an amateur investor, I am often interested in the Compound Annual
Growth Rate (CAGR) calculation for comparing the smoothed rate of return
of different investments. Surprisingly enough, Excel doesn't have this
in its kitbag, so I wrote my own. Below is my version of a user-defined
function (UDF) that can be used in Excel.

```vb
Public Function CAGR( _
   ByVal StartValue As Double, _
   ByVal EndValue As Double, _
   ByVal StartDate As Date, _
   ByVal EndDate As Date) _
   As Double

   ' Compute Compound Annual Growth Rate according to formula
   ' CAGR = (FV / PV ) ^ 1/n - 1 where n is number of years
   ' Developed by Warren Bain of Thought Croft Pty Ltd

   CAGR = (EndValue / StartValue) _
          ^ (1 / ((EndDate - StartDate) / 365.25)) - 1
End Function
```
