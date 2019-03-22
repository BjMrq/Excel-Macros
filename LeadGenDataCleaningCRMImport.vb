Sub facebookLeadImportToCRMOpenHouse()
'
' facebookLeadImportToCRMOpenHouse Macro
' Clean and consolidate the data from Facebook Lead Gen before importing into the CRM
'

    eraseUnecesaryColumns
    formatHeader
    formatOptInAndPhone
    ExpOptInDate
    implicitOptIn
    implicitOptInDate
    rating
    source
    medium
    content
    firstContact
    marketingEvent
    leadSource
    campus
    eventDetail
    language
    endFormat
    saveLeadGenImportReady



End Sub

Sub eraseUnecesaryColumns()

    Columns("A:L").Select
    Selection.Delete Shift:=xlToLeft

End Sub

Sub formatHeader()


    Range("A1").Select
    ActiveCell.FormulaR1C1 = "E-mail 1"

    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Home Phone"

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "First Name"

    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Last Name"

    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Explicit Opt-in"

    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Explicit Opt-in Date"

    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Implicit Opt-in"

    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Implicit Opt-in Date"

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Lead Source (Marketo)"

    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Lead Source Detail"

    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Rating"

    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Source"

    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Medium"

    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Content"

    Range("O1").Select
    ActiveCell.FormulaR1C1 = "First Contact"

    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Campus"

    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Event Source"

    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Marketing Event"

    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Language"


End Sub

Sub formatOptInAndPhone()


   Cells.Replace What:="p:+", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="TRUE", Replacement:="yes", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="False", Replacement:="no", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


End Sub

Sub ExpOptInDate()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("F2").Select

    Dim thisDate As String
    thisDate = Date


    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("E" & i).Value = "yes" Then

        ActiveCell.Value = thisDate

    End If

        ActiveCell.Offset(1, 0).Select



    Next i


End Sub

Sub implicitOptIn()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("G2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("E" & i).Value = "no" Then

        ActiveCell.Value = "yes"

        ElseIf Range("E" & i).Value = "yes" Then

        ActiveCell.Value = "no"

    End If

        ActiveCell.Offset(1, 0).Select



    Next i


End Sub

Sub implicitOptInDate()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("H2").Select

    Dim thisDate As String
    thisDate = Date


    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("E" & i).Value = "no" Then

        ActiveCell.Value = thisDate

    End If

        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub rating()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("K2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "qualify warm (+50%)"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub source()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("L2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "Facebook"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub medium()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("M2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "Banner"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub content()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("N2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "Lead Gen"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub firstContact()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("O2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "Marketing Event"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub marketingEvent()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("R2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "Registered"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub leadSource()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("I2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "event"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub campus()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Dim campus As String

    Message = "For What Campus is your Lead Gen?"
    Title = "campusInput"

    campus = InputBox(Message, Title)

    Range("P2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = campus


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub eventDetail()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count

    Dim detail As String

    Message = "What event want to attend this Lead Gen?"
    Title = "detailInput"

    detail = InputBox(Message, Title)

    Range("J2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = detail


        ActiveCell.Offset(1, 0).Select

    Next i


    Range("Q2").Select

    Dim j As Integer

    For j = 2 To nbRecords + 1

        ActiveCell.Value = detail


        ActiveCell.Offset(1, 0).Select


    Next j



End Sub
Sub language()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Dim language As String

    Message = "What is the langage of this Lead Gen?"
    Title = "languageInput"

    language = InputBox(Message, Title)

    Range("S2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = language


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub
Sub endFormat()

    Columns("A:S").Select
    Selection.ColumnWidth = 12

    Columns("G:G").ColumnWidth = 6
    Columns("E:E").ColumnWidth = 6
    Columns("J:J").ColumnWidth = 26.43
    Columns("O:O").ColumnWidth = 17.14
    Columns("Q:Q").ColumnWidth = 26.43
    Columns("H:H").ColumnWidth = 10.29
    Columns("F:F").ColumnWidth = 11.14
    Columns("M:M").ColumnWidth = 9.57
    Columns("N:N").ColumnWidth = 10
    Columns("O:O").ColumnWidth = 15.86
    Columns("L:L").ColumnWidth = 10.71
    Columns("K:K").ColumnWidth = 18.86
    Columns("I:I").ColumnWidth = 9.43

    Range("B2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        Selection.NumberFormat = "0;00"

        ActiveCell.Offset(1, 0).Select


    Next i



End Sub

Sub saveLeadGenImportReady()

    ActiveWorkbook.SaveAs Filename:= _
          "C:\Users\bmarquis\OneDrive - College LaSalle\Downloads\cleanLeadGenReadyToImport.csv", FileFormat _
          :=xlCSVUTF8, CreateBackup:=False

End Sub
