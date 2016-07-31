---
layout: post
title:  "Generic function to copy Excel data between workbooks"
date:   2007-10-12 15:33:00 +1000
tags:   [excel, workbook, vba]
---

I needed way to ensure that some Excel workbooks that were being
completed by customers and uploaded to a server for analysis and
subsequent loading into a database, were not being tampered with.
Certainly in Excel version up to and including 2003 it is not possible
to prevent a determined user from "cracking" the Worksheet and Workbook
protection passwords and hence to modifying any formulae in the
Workbook.

My solution was to create a new, clean copy of the workbook from a
master template and to then copy the customer entered data from the
submitted book to the new clean book. The following CopyExcelData sub
does the bulk of the work - the comments explain what is going on.

It does take some time to do this as it operates cell-by-cell (around 1
minute to copy 500 cells spread across 15 worksheets) but its not a bad
trade-off to be safe in the knowledge that the final result is
tamper-proof.

As a side benefit, any corrections needed to the customer workbook after
it has been issued can be made to the template version and files
received after that point will be automatically "upgraded".

```vb
Public Sub CopyExcelData( _
       ByRef wkbSource As Object, _
       ByRef wkbTarget As Object, _
       Optional ByVal blnCopyEmptyCells As Boolean = True)

    '*** Change to remove control chars as it crashes Excel 97 ***'

    ' Copy all data entry cells from one workbook
    ' to the other assuming that a data entry cell
    ' is:
    '   1) On Visible sheets only
    '   2) In the UsedRange of cells
    '   3) If Sheet is Protected then Unlocked Cells
    '   4) If Sheet is UnProtected then Non-formula cells
    '
    ' Since the target is expected to be the 'good' copy,
    ' that is the one we use to test the above conditions
    ' and we then extract the corresponding data from the
    ' source cell and place it in the target cell

    ' Note: late binding has been used to limit any issues
    ' related to different versions of Excel, parameters
    ' are actually:
    '               ByRef wkbSource As Excel.Workbook
    '               ByRef wkbTarget As Excel.Workbook

    Dim appExcel              As Object          'Excel.Application
    Dim blnProtectTarget      As Boolean
    Dim rngAllTarget          As Object          'Excel.Range
    Dim rngCellSource         As Object          'Excel.Range
    Dim rngCellTarget         As Object          'Excel.Range
    Dim wksSource             As Object          'Excel.Worksheet
    Dim wksTarget             As Object          'Excel.Worksheet
    Dim xlCalcMode            As Variant

    ' Before we start, ensure calculation mode is manual
    Set appExcel = wkbSource.Application
    xlCalcMode = appExcel.Calculation
    appExcel.Calculation = xlCalculationManual

    For Each wksTarget In wkbTarget.Worksheets
        If wksTarget.Visible = xlSheetVisible Then
            ' We only want data on sheets the user can see
            ' so we ignore any that are Hidden or VeryHidden
            Set rngAllTarget = wksTarget.UsedRange
            If Not rngAllTarget Is Nothing Then
                ' We have some non-empty cells on this sheet
                Set wksSource = wkbSource.Worksheets(wksTarget.Name)
                blnProtectTarget = wksTarget.ProtectContents

                For Each rngCellTarget In rngAllTarget.Cells
                    ' Stepping through each cell in the range...
                    With rngCellTarget
                        If (blnProtectTarget And Not .Locked) Or _
                           (Not blnProtectTarget And Not .HasFormula) Then
                            ' This is a cell that can be completed in
                            ' the original target sheet so examine further
                            If .Address = .MergeArea(1, 1).Address Then
                                ' This is the main cell for a merged set of cells
                                ' or not merged at all so we are interested...
                                Set rngCellSource = wksSource.Range(.Address)
                                If Not IsError(rngCellSource.Value2) Then
                                    ' Only copy valid cell entries
                                    If rngCellSource.HasFormula And _
                                       Not (rngCellSource.FormulaHidden Or .FormulaHidden) Then
                                        ' They are using a formula and we can access the formula
                                        ' in both source and target so transfer it (can't access this
                                        ' property if FormulaHidden is TRUE for either)
                                        .Formula = rngCellSource.Formula
                                    ElseIf Len(rngCellSource.Value2) > 0 Or blnCopyEmptyCells Then
                                        ' Not a formula so just get the value using Value2
                                        ' to avoid problems introduced by incorrect date formats
                                        ' NOTE: remove control characters to avoid Excel 97 crash
                                        .Value2 = tcStripChars(rngCellSource.Value2, scmcRemoveControl)
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next rngCellTarget
            End If
        End If
    Next wksTarget

    ' Return calculation mode to whatever it was before
    appExcel.Calculation = xlCalcMode

    Set rngCellTarget = Nothing
    Set rngCellSource = Nothing
    Set wksSource = Nothing
    Set wksTarget = Nothing
    Set appExcel = Nothing
End Sub

Public Function GetNamedRangeValue(ByRef nm As Object) As Variant

    ' To get the value held by a range name.  This
    ' function handles Named constants and formulae
    ' which can't be evaluated by the object itself

    ' Note: to avoid problems with different Excel
    ' versions, we use late binding of the range
    ' and the input parameter:
    '       ByRef nm As Excel.Name
    '       Dim rng As Excel.Range

    Dim rng                   As Object          ' Excel.Range

    With nm
        ' Check to see if this is a named constant or formula
        ' in which case it won't have a range object
        On Error Resume Next
        Set rng = .RefersToRange
        On Error GoTo 0
        If rng Is Nothing Then
            ' This a named constant or named formula
            ' so we need to use Excel to evaluate
            On Error Resume Next
            GetNamedRangeValue = .Application.ExecuteExcel4Macro(Mid(.RefersToR1C1, 2))
            On Error GoTo 0
        Else
            ' This is a cell so we can recover the value
            ' using the RefersToRange value2 which allows
            ' us better control over formatting glitches
            GetNamedRangeValue = .RefersToRange.Value2
        End If
    End With
    Set rng = Nothing
End Function
```
