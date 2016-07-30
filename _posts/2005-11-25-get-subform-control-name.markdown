---
layout: post
title:  "Finding the name of a control in an Access subform"
date:   2005-11-25 12:06:00 +1000
tags:   [access, form, subform, vba]
---

I'm currently developing a Microsoft Access based system and found
myself continually needing to work out how to access a property of the
control that contains a subform from code running in the subform (for
example to get at the Tag property).

This function will do that by walking the controls collection of the
subform's parent form looking for any subforms and then compares the
hWnd (basically Windows internal "handle" for that window) of that
control with the hWnd of our subform.

Once found, we construct the name of the control using the appropriate
name format. If we want to use the name in code then the short format is
fine but if it is to be used in a query then we need the long version
which may necessitate walking up the hierarchy if in fact the parent
form is itself a subform (forms can be nested to three levels). This is
achieved by calling the function recursively on the parent.

```vb
Public Enum ControlNameFormat
   cnfcShortPropertyName
   cnfcLongHierarchicalName
End Enum

Public Function GetSubFormControlName( _
   ByRef frm As Form, _
   Optional ByVal NameFormat As ControlNameFormat = _
      cnfcShortPropertyName) As String

' Tells a subform the name of the control that
' it has been opened in on the main form. Used
' to modify the base subform's source or to
' retrieve special values from its Tag property.
' The NameFormat tells us whether to just provide
' the ctl.Name property or the fully qualified
' form controls collection item name.

' Example:
' Form "MainForm" holding "1stSubForm" in control "fsub1"
' holding "2ndSubForm" in control "fsub2"...
' GSFCN(1stSubForm,Long) = "[Forms]![MainForm]![fsub1]"
' GSFCN(2ndSubForm,Long) = "[Forms]![MainForm]![fsub1].Form![fsub2]"
' GSFCN(2ndSubForm,Short) = "fsub2"

' Developed by Warren Bain on 26/09/2005
' Copyright (c) Thought Croft Pty Ltd.

   Dim ctl As Control
   Dim strResult As String

   On Error Resume Next

   ' Loop through all controls on the parent and test for the
   ' handle of the window of the subsidiary form. If it
   ' matches ours, then we have found the control it opened in
   If Not IsSubForm(frm) Then
      ' We are at the top of the tree, so return name
      If NameFormat = cnfcShortPropertyName Then
         strResult = frm.Name
      ElseIf NameFormat = cnfcLongHierarchicalName Then
         strResult = "[Forms]![" & frm.Name & "]"
      End If
   Else
      For Each ctl In frm.Parent.Controls
         If ctl.ControlType = acSubform Then
            If ctl.Form.hWnd = frm.hWnd Then
               ' Found the right one
               If NameFormat = cnfcShortPropertyName Then
                  ' Just return the name of the control
                  strResult = ctl.Name
               ElseIf NameFormat = cnfcLongHierarchicalName Then
                  ' Add parent plus fully qualified control
                  strResult = GetSubFormControlName(frm.Parent, NameFormat) & ".Form![" & ctl.Name & "]"
               End If
               Exit For
            End If
         End If
      Next ctl
   End If
   GetSubFormControlName = strResult
End Function

Private Function IsSubForm(frm As Form) As Boolean
   ' Is the form currently loaded as a subform?
   Dim strFormName As String
   On Error Resume Next
   strFormName = frm.Parent.Name
   IsSubForm = (Err.Number = 0)
   Err.Clear
End Function
```
