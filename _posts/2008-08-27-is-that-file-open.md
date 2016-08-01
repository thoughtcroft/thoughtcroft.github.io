---
layout: post
title:  "Is that file open"
date:   2008-08-27
tags:   [file, windows, vba]
---

f you ever need to determine if a file is already open before you attempt to do something with it in VBA, you will discover a blistering array of code segments on the interweb which all seem to share the characteristic that: **they fail to work as advertised**.

So time for me to add my version and I can state that it absolutely works except where the opening application fails to lock the file. AFAIK, that comment applies to text files opened in Notepad and any file opened with an Office application that has its read-only attribute set. However, since the main reason for testing if a file is already open is so that you can do stuff to it without it causing a "file is already in use" type of error, these two conditions don't matter since they don't stop you doing anything with the file. If you don't believe me, try this:

* Create an Excel file "test.xls"
* In Windows Explorer right click on the file and mark it as "read-only"
* Open the file in Excel
* Now go back to Windows Explorer and delete the file

Voila! No problem since it was read-only, Excel opened a temporary copy of the file and so you do stuff to the original without it causing any problems.

Now, back to the problem: how do you tell if a given file is in use?
Answer: use the VBA "open" statement to exclusively open the file and trap any errors that may occur. We have to go one further though as the "open" statement will not see any hidden files and will act as though they don't exist. To overcome this we make some attribute changes and trap any errors that may occur in relation to them as well.

```vb
Public Function IsFileOpen(ByVal strFullPathFileName As String) As Boolean

   ' Attempting to open a file for ReadWrite that exists will fail
   ' if someone else has it open.  We also have to guard against the
   ' errors that occur if the file has uncommon file attributes such as
   ' 'hidden' which can upset the Open statement.
   ' NOTE: any open that doesn't lock the file such as opening a .txt file
   ' in NotePad or a read-only file open will return False from this call.

   Dim lngFile                    As Long
   Dim intAttrib                  As Integer

   On Error Resume Next
   intAttrib = GetAttr(strFullPathFileName)
   If Err <> 0 Then
       ' If we can't get these then it means the file name is
       ' invalid, or the file or path don't exist so no problem
       IsFileOpen = False
       Exit Function
   End If

   SetAttr strFullPathFileName, vbNormal
   If Err <> 0 Then
       ' An error here means that the file is open and the attributes
       ' therefore can't be changed so let them know that
       IsFileOpen = True
       Exit Function
   End If

   ' Ready to try and open the file exclusively and then any error means that
   ' the file is already open by some other process...
   lngFile = FreeFile
   Open strFullPathFileName For Random Access Read Write Lock Read Write As lngFile
   IsFileOpen = (Err <> 0)
   Close lngFile

   ' Restore the attributes and exit
   SetAttr strFullPathFileName, intAttrib

End Function
```
