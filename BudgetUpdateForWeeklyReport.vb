Sub updateBudget()

    copyLeadsIntoPreviousWeek
    copySpentIntoPreviousWeek
    updateData
    copySpentIntoWeeklyBudgetPlanning


End Sub
Sub copyLeadsIntoPreviousWeek()
'
'   Copy paste Leads

    Range("C16").Select
    Selection.Copy
    Range("C20").Select
    ActiveSheet.Paste

    Range("C17").Select
    Selection.Copy
    Range("C21").Select
    ActiveSheet.Paste

    Range("C53").Select
    Selection.Copy
    Range("C57").Select
    ActiveSheet.Paste

    Range("C54").Select
    Selection.Copy
    Range("C58").Select
    ActiveSheet.Paste

    Range("C91").Select
    Selection.Copy
    Range("C95").Select
    ActiveSheet.Paste

    Range("C92").Select
    Selection.Copy
    Range("C96").Select
    ActiveSheet.Paste




End Sub
Sub copySpentIntoPreviousWeek()
'
' Macro5 Macro
'
    Range("E16").Copy
    Range("E20").PasteSpecial xlPasteValues

    Range("E17").Copy
    Range("E21").PasteSpecial xlPasteValues

    Range("E53").Copy
    Range("E57").PasteSpecial xlPasteValues

    Range("E54").Copy
    Range("E58").PasteSpecial xlPasteValues

    Range("E91").Copy
    Range("E95").PasteSpecial xlPasteValues

    Range("E92").Copy
    Range("E96").PasteSpecial xlPasteValues


End Sub

Sub copySpentIntoWeeklyBudgetPlanning()

    Range("E16").Copy
    Range("C5").Select
    Dim i As Integer
    For i = 1 To 12
        If ActiveCell.Value = "" Then
            ActiveCell.PasteSpecial xlPasteValues
            Exit For
        Else
            ActiveCell.Offset(0, 1).Select
        End If
    Next i


    Range("E53").Copy
    Range("C42").Select
    Dim j As Integer
    For j = 1 To 12
        If ActiveCell.Value = "" Then
            ActiveCell.PasteSpecial xlPasteValues
            Exit For
        Else
            ActiveCell.Offset(0, 1).Select
        End If
    Next j


    Range("E91").Copy
    Range("C80").Select
    Dim k As Integer
    For k = 1 To 12
        If ActiveCell.Value = "" Then
            ActiveCell.PasteSpecial xlPasteValues
            Exit For
        Else
            ActiveCell.Offset(0, 1).Select
        End If
    Next k


End Sub

Sub updateData()

    ActiveWorkbook.RefreshAll

End Sub

Sub UpdateFacebookBudget()
'
' UpdateFacebookBudget Macro
'

'Calculate Inter-Dec budget
    Sheets("FacebookLastWeek").Select

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("A2").Select

    Dim interDecSpent As Integer

    interDecSpent = 0

    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("A" & i).Value Like "CID*" Then

        interDecSpent = interDecSpent + Range("D" & i).Value

    End If

        ActiveCell.Offset(1, 0).Select

    Next i

    Range("I2").Value = interDecSpent


    'Calculate LaSalle budget

    Range("A2").Select

    Dim laSalleSpent As Integer

    laSalleSpent = 0

    Dim j As Integer

    For j = 2 To nbRecords + 1

    If Range("A" & j).Value Like "CLM*" Then

        laSalleSpent = laSalleSpent + Range("D" & j).Value

    End If

        ActiveCell.Offset(1, 0).Select



    Next j

    Range("I3").Value = laSalleSpent


    'Calculate eLearning budget

    Range("A2").Select

    Dim eLearningSpent As Integer

    eLearningSpent = 0

    Dim k As Integer

    For k = 2 To nbRecords + 1

    If Range("A" & k).Value Like "eLearning*" Then

        eLearningSpent = eLearningSpent + Range("D" & k).Value

    End If

        ActiveCell.Offset(1, 0).Select



    Next k

    Range("I4").Value = eLearningSpent

End Sub
