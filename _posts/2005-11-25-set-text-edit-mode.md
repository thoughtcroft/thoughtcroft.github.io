---
layout: post
title:  "Setting text edit mode on Access form fields made easy"
date:   2005-11-25 15:34:00 +1000
tags:   [access, control, form, textfield, vba]
---

The Locked and Enabled properties of text-based controls - combo boxes, check boxes, list boxes and text boxes
- control whether they can be changed or entered. But for the life of me, I can never remember which combination
of true and false values gives me the look I am after. For example, `Enabled=Yes` and `Locked=Yes` means the text
field can be entered but can't be changed. Change this to `Enabled=No` and `Locked=Yes` and you won't be able to
enter or edit the field. Make it `Enabled=Yes`, `Locked=No` and you can enter and edit and so on.

To make it easier on myself, I wrote this function to do the remembering for me. As you can see, a lot of the work
is done by the enumerated constants definition - that's why I like to use them apart from the fact that the
VBA editor also reminds me what values I can use when I am coding. I also like to use these constants as bitwise
comparison flags by making them different powers of 2 - easier to look at the example than try and explain!

The net effect is that some sensible constant design simplifies coding to two statements rather than a string of nested IFs.

```vb
Public Enum TextEditMode
   temcInvalid = 0
   temcLockedTrue = 2 ^ 1
   temcLockedFalse = 2 ^ 2
   temcEnabledTrue = 2 ^ 3
   temcEnabledFalse = 2 ^ 4
   temcEnterWithEdit = temcLockedFalse + temcEnabledTrue
   temcEnterNoEdit = temcLockedTrue + temcEnabledTrue
   temcNoEnterNormal = temcLockedTrue + temcEnabledFalse
   temcNoEnterDimmed = temcLockedFalse + temcEnabledFalse
End Enum

Public Function SetTextEditMode( _
   ByRef ctl As Control, _
   Optional ByVal temMode As TextEditMode) As TextEditMode

   ' Set or return the value of Enabled and Locked properties
   ' in a text based control to manage how it looks as follows:
   ' Enabled? Locked? Result?
   ' Yes Yes Can enter, can't edit, normal
   ' Yes No Can enter, can edit, normal
   ' No Yes Can't enter, can't edit, normal
   ' No No Can't enter, can't edit, dimmed
   ' If no mode requested then returns the current settings

   ' Developed by Warren Bain on 21/10/2005
   ' Copyright (c) Thought Croft Pty Ltd
   ' All rights reserved.

   ' Check we can do this for this type of control
   Select Case ctl.ControlType
      Case acComboBox, acCheckBox, acListBox, acTextBox
         If IsMissing(temMode) Then
            ' Let them know what is set
            SetTextEditMode = IIf(ctl.Enabled, temcEnabledTrue, temcEnabledFalse) + _
               IIf(ctl.Locked, temcLockedTrue, temcLockedFalse)
         Else
            ' Set the controls parameters
            ctl.Enabled = temMode And temcEnabledTrue
            ctl.Locked = temMode And temcLockedTrue
            SetTextEditMode = temMode
         End If
      Case Else
         SetTextEditMode = temcInvalid
      End Select
End Function
```
