Sub A___UpdateWeeklyReport()
'
'
    copyLeadsIntoPreviousWeek
    copySpentIntoPreviousWeek


End Sub
Sub A___RefreshWeeklyData()
'
'
    refreshData
    updateFacebookBudget


End Sub

Sub copyLeadsIntoPreviousWeek()
'
'   Copy paste Leads

    Sheets("Weekly Stats").Select

    Range("C5").Select
    Selection.Copy
    Range("C9").Select
    Selection.PasteSpecial xlPasteValues

    Range("C6").Select
    Selection.Copy
    Range("C10").Select
    Selection.PasteSpecial xlPasteValues

    Range("C18").Select
    Selection.Copy
    Range("C22").Select
    Selection.PasteSpecial xlPasteValues

    Range("C19").Select
    Selection.Copy
    Range("C23").Select
    Selection.PasteSpecial xlPasteValues

    Range("C31").Select
    Selection.Copy
    Range("C35").Select
    Selection.PasteSpecial xlPasteValues

    Range("C32").Select
    Selection.Copy
    Range("C36").Select
    Selection.PasteSpecial xlPasteValues


End Sub
Sub copySpentIntoPreviousWeek()

    Range("E5").Select
    Selection.Copy
    Range("E9").Select
    Selection.PasteSpecial xlPasteValues

    Range("E6").Select
    Selection.Copy
    Range("E10").Select
    Selection.PasteSpecial xlPasteValues

    Range("E18").Select
    Selection.Copy
    Range("E22").Select
    Selection.PasteSpecial xlPasteValues

    Range("E19").Select
    Selection.Copy
    Range("E23").Select
    Selection.PasteSpecial xlPasteValues

    Range("E31").Select
    Selection.Copy
    Range("C35").Select
    Selection.PasteSpecial xlPasteValues

    Range("E32").Select
    Selection.Copy
    Range("E36").Select
    Selection.PasteSpecial xlPasteValues

End Sub

Sub updateFacebookBudget()
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


Sub refreshData()
'
' refresh Macro
'

'
    ActiveWorkbook.RefreshAll
    Sheets("Sheet1").Select
    Range("E5").Select
    With ActiveSheet.PivotTables("PivotTable3").PivotFields( _
        "Latest Campus of Interest")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveWindow.SmallScroll Down:=-21
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("First Name"), "Count of First Name", xlCount
    ActiveWindow.SmallScroll Down:=6
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("GA Source")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("GA Medium")
        .Orientation = xlColumnField
        .Position = 2
    End With
    Sheets("Weekly Stats").Select

End Sub
