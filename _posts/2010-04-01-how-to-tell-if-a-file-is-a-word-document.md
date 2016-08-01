---
layout: post
title:  "How to tell if a file is a Word document"
date:   2010-04-01
tags:   [file, word, vba]
---

I am building a tool that needs to do something to all the Word documents in a set of folders. There are other sorts of files in these folders so I need to filter the list. I didn't want to hard code the file extensions that constitute Word documents because these change according to the version and will do so in future. So I decided to develop a function that is cross-version compatible.

To determine the documents that Microsoft Word can open and update natively, I am using the `Filters` collection from the `Application.FileDialog` object which is available in Word version 2002 onwards. This is the list of extensions that can be selected in the File Open dialog to filter the files that can be opened in Word. Seems to work like a charm. For earlier versions, you'll need to hardcode the extensions.

```vb
Public Function IsWordDocument(ByVal strExtension As String) As Boolean

    ' Developed by Warren Bain on 01/04/2010
    ' Copyright (c) Thought Croft Pty Ltd
    ' All rights reserved.

    ' Verifies if the supplied file extension e.g. "doc"
    ' is recognised as one of the documents that this version
    ' of Word can natively handle.  List is constructed from the
    ' types of documents that can be filtered in the Open File dialog
    ' therefore this only works from Microsoft Word 2002 onwards.

    Const strcNoiseChars           As String = "*?."
    Const strcSeparators           As String = ";,"
    Const strcDelimiter            As String = "|"

    Static colExtensions           As Collection

    Dim fdf                        As FileDialogFilter
    Dim astrExts                   As Variant
    Dim strExts                    As String
    Dim i                          As Integer

    ' Check if we have loaded the collection yet - only done once
    If colExtensions Is Nothing Then
        Set colExtensions = New Collection
        For Each fdf In Application.FileDialog(msoFileDialogOpen).Filters
            strExts = fdf.Extensions

            ' Remove any 'noise' characters from the string
            For i = 1 To Len(strcNoiseChars)
                strExts = Replace(strExts, Mid$(strcNoiseChars, i, 1), vbNullString)
            Next i

            ' Ensure we standardise on separators used
            For i = 1 To Len(strcSeparators)
                strExts = Replace(strExts, Mid$(strcSeparators, i, 1), strcDelimiter)
            Next i

            ' Turn the current set of extensions into an array
            astrExts = Split(strExts, strcDelimiter)

            ' Add all the ones we haven't already got
            For i = LBound(astrExts) To UBound(astrExts)
                ' If already there, this will fail so ignore
                On Error Resume Next
                colExtensions.Add Trim(astrExts(i)), Trim(astrExts(i))
                On Error GoTo 0
            Next i
        Next fdf
    End If

    ' We just try and look up the file type and if it fails
    ' then it is not an intrinsically supported document
    On Error Resume Next
    strExts = colExtensions.Item(strExtension)
    IsWordDocument = (Err = 0)
End Function
```
