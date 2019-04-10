Dim detail As String ' Module-level variable.

Sub A___facebookLeadImportELEARNING()
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
    geoCountry
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
    ActiveCell.FormulaR1C1 = "Program Choice 1"

    Range("B1").Select
    ActiveCell.FormulaR1C1 = "E-mail 1"

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Home Phone"

    Range("D1").Select
    ActiveCell.FormulaR1C1 = "First Name"

    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Last Name"

    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Explicit Opt-in"

    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Explicit Opt-in Date"

    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Implicit Opt-in"

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Implicit Opt-in Date"

    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Lead Source (Marketo)"

    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Lead Source Detail"

    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Rating"

    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Source"

    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Medium"

    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Content"

    Range("P1").Select
    ActiveCell.FormulaR1C1 = "First Contact"

    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Campus"

    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Event Source"

    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Marketing Event"

    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Language"

    Range("U1").Select
    ActiveCell.FormulaR1C1 = "GeoCountry"


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

    Cells.Replace What:="e-business", Replacement:="LEACE-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="affaires_électroniques", Replacement:="LEACE-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="administrative_assistant", Replacement:="LCE6S-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="adjoint_administratif", Replacement:="LCE6S-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="event_planning_and_management", Replacement:="LCAD0-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="planification_et_gestion_d'événements", Replacement:="LCAD0-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="stylisme_de_mode", Replacement:="NTC0L-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="fashion_styling", Replacement:="NTC0L-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="modélisation_3d_de_jeux_vidéo", Replacement:="NTL0Y-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="video_game_3d_modeling", Replacement:="NTL0Y-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="design_infographique", Replacement:="NWC0W-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="infographic_design", Replacement:="NWC0W-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="commercialisation_de_la_mode", Replacement:="NTC1H-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="fashion_marketing", Replacement:="NTC1H-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="design_d'intérieur", Replacement:="NTA1P-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="interior_design", Replacement:="NTA1P-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="intégration_multimédia", Replacement:="NWE30-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="multimedia_integration", Replacement:="NWE30-elarning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("C2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        Selection.NumberFormat = "0;00"


        ActiveCell.Offset(1, 0).Select


    Next i


End Sub

Sub ExpOptInDate()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("G2").Select

    Dim thisDate As String
    thisDate = Date


    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("F" & i).Value = "yes" Then

        ActiveCell.Value = thisDate

    End If

        ActiveCell.Offset(1, 0).Select



    Next i


End Sub

Sub implicitOptIn()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("H2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("F" & i).Value = "no" Then

        ActiveCell.Value = "yes"

        ElseIf Range("F" & i).Value = "yes" Then

        ActiveCell.Value = "no"

    End If

        ActiveCell.Offset(1, 0).Select



    Next i


End Sub

Sub implicitOptInDate()

    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("I2").Select

    Dim thisDate As String
    thisDate = Date


    Dim i As Integer

    For i = 2 To nbRecords + 1

    If Range("F" & i).Value = "no" Then

        ActiveCell.Value = thisDate

    End If

        ActiveCell.Offset(1, 0).Select


    Next i

End Sub

Sub rating()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("L2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "qualify warm (+50%)"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub
Sub geoCountry()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("U2").Select


    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = "Canada"


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub
Sub source()


    Dim nbRecords As Integer

    Range("A2").Select
    nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count


    Range("M2").Select


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


    Range("N2").Select


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


    Range("O2").Select


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


    Range("P2").Select


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


    Range("S2").Select


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


    Range("J2").Select

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

    Range("Q2").Select

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



    Message = "What event want to attend this Lead Gen?"
    Title = "detailInput"

    detail = InputBox(Message, Title)

    Range("K2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = detail


        ActiveCell.Offset(1, 0).Select

    Next i


    Range("R2").Select

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

    Range("T2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        ActiveCell.Value = language


        ActiveCell.Offset(1, 0).Select


    Next i

End Sub
Sub endFormat()



    Range("C2").Select

    Dim i As Integer

    For i = 2 To nbRecords + 1

        Selection.NumberFormat = "0;00"

        ActiveCell.Offset(1, 0).Select


    Next i



End Sub

Sub saveLeadGenImportReady()

    ActiveWorkbook.SaveAs Filename:= _
          "C:\Users\bmarquis\OneDrive - College LaSalle\Downloads\" & detail & ".csv", FileFormat _
          :=xlCSVUTF8, CreateBackup:=False

End Sub
