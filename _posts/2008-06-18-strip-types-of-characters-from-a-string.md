---
layout: post
title:  "Strip different types of characters from a string"
date:   2008-06-18
tags:   [excel, string, vba]
---

I realised that one of my earlier posts Generic Function to Copy Excel
Data references a function called 'tcStripChars' in order to remove
control characters from a cell value. This was required to prevent Excel
97 crashing when copying the .Value2 property of a cell into another
cell whenever there were control characters in the value e.g. LFCR.

The function is useful for other situations as well so here it is.
Please note the use of the Vba6 conditional compilation constant to
generate appropriate code for your version of VBA using Enumerated Types
(VBA6) or Public Constants (pre VBA6).

```vb
' Note: Vba6 is a conditional compiler constant that indicates
' the version of VBA and in this module we use enumerated types if
' supported otherwise we use plain old public constants

#If CBool(VBA6) Then
' Enumerate methods for selecting mode of character removal
Public Enum StripCharsMode
    scmcRemoveAlphas = 2 ^ 0
    scmcRemoveControl = 2 ^ 1
    scmcRemoveNumerics = 2 ^ 2
    scmcRemoveSpaces = 2 ^ 3
    scmcRemoveOthers = 2 ^ 4
    scmcRemoveAll = 2 ^ 5 - 1
    scmcKeepAlphas = scmcRemoveAll - scmcRemoveAlphas
    scmcKeepLetters = scmcRemoveAll - scmcRemoveAlphas - scmcRemoveSpaces
    scmcKeepNumerics = scmcRemoveAll - scmcRemoveNumerics
    scmcKeepOthers = scmcRemoveAll - scmcRemoveOthers
    scmcKeepControl = scmcRemoveAll - scmcRemoveControl
End Enum

Public Function tcStripChars( _
       ByVal strInputText As String, _
       ByVal scmRemoveType As StripCharsMode) _
       As String

    Dim scmCharMode As StripCharsMode
#Else
' Constants for selecting mode of character removal
Public Const scmcRemoveAlphas As Integer = 2 ^ 0
Public Const scmcRemoveControl As Integer = 2 ^ 1
Public Const scmcRemoveNumerics As Integer = 2 ^ 2
Public Const scmcRemoveSpaces As Integer = 2 ^ 3
Public Const scmcRemoveOthers As Integer = 2 ^ 4
Public Const scmcRemoveAll As Integer = 2 ^ 5 - 1
Public Const scmcKeepAlphas As Integer = scmcRemoveAll - scmcRemoveAlphas
Public Const scmcKeepLetters As Integer = scmcRemoveAll - scmcRemoveAlphas _
                                          - scmcRemoveSpaces
Public Const scmcKeepNumerics As Integer = scmcRemoveAll - scmcRemoveNumerics
Public Const scmcKeepOthers As Integer = scmcRemoveAll - scmcRemoveOthers
Public Const scmcKeepControl As Integer = scmcRemoveAll - scmcRemoveControl

Public Function tcStripChars( _
       ByVal strInputText As String, _
       ByVal scmRemoveType As Integer) _
       As String

    Dim scmCharMode As Integer
#End If

    Dim intPos As Integer
    Dim strChar As String

    ' Remove specified types of characters from input string

    ' Developed by Warren Bain
    ' Copyright 2004, Thought Croft Pty Ltd
    ' All rights reserved.

    ' In:
    '   strInputText:
    '       text to extract characters from
    '   scmRemoveType:
    '       type of removal (or retention) required
    '       can be combined to remove multiple types of chars
    ' Out:
    '   Return Value:
    '       text with all required chars removed

    ' Start with an empty output string
    tcStripChars = vbNullString

    ' Determine for each character in the input string
    ' whether it should be retained or discarded based
    ' on bitwise comparison with scmRemoveType parameter
    For intPos = 1 To Len(strInputText)
        strChar = Mid$(strInputText, intPos, 1)
        Select Case Asc(strChar)
        Case 65 To 90, 97 To 122: scmCharMode = scmcRemoveAlphas
        Case 48 To 57: scmCharMode = scmcRemoveNumerics
        Case 32: scmCharMode = scmcRemoveSpaces
        Case 0 To 31: scmCharMode = scmcRemoveControl
        Case Else: scmCharMode = scmcRemoveOthers
        End Select

        ' If the character's type bit is set in the remove type
        ' we will discard it - otherwise retain it
        If scmRemoveType And scmCharMode Then
            'Ignore this one
        Else
            tcStripChars = tcStripChars & strChar
        End If
    Next intPos
End Function
```
