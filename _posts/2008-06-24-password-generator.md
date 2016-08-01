---
layout: post
title:  "Password generator"
date:   2008-06-24
tags:   [password, vba]
---

This is a simple digit only password generator that I use for a number
of purposes. My recent post about managing automation objects uses this
function and I forgot to include it there - I have updated that post but
though I would also post it separately. I adapted this from some code I
got somewhere, my apologies for not acknowledging the source.

```vb
Public Function GeneratePassword( _
       ByVal intLength As Integer) As String

    ' Generates a random string of digits of the requested length

    ' In:
    '   intLength - number of digits to be returned (max 9)
    ' Out:
    '   Return Value - a random string of digits
    ' Example:
    '   GetPassword(3) = "927"

    Dim lngHighNumber              As Long
    Dim lngLowNumber               As Long
    Dim lngRndNumber               As Long

    ' Check we don't exceed our maximum range
    If intLength > 9 Or intLength < 1 Then
        Err.Raise 5, "GetPassword", _
                  "Invalid string length - must be between 1 and 9"
    Else
        ' Work out the numbers
        lngLowNumber = 10 ^ (intLength - 1)
        lngHighNumber = (10 ^ intLength) - 1
        ' Generate a new seed and a new random number
        Randomize
        lngRndNumber = Int((lngHighNumber - lngLowNumber + 1) * Rnd) + lngLowNumber
        ' Format the result as string
        GeneratePassword = Format$(lngRndNumber, String$(intLength, "0"))
    End If
End Function
```
