Sub A___updateCampaignReport()

    Sheets("PO Management").Select
    Dim nbPO As Integer

    Range("T5").Select
    nbPO = Range(Selection, Selection.End(xlDown)).Rows.Count

    Dim i As Integer

    For i = 5 To nbPO + 4
        Range("T" & i).Select
        PO = ActiveCell.Value

        Sheets("FacebookSpending").Select

        Dim nbRecords As Integer

        Range("A2").Select
        nbRecords = Range(Selection, Selection.End(xlDown)).Rows.Count

        Dim POSpent As Integer

        POSpent = 0

        Dim j As Integer

            For j = 2 To nbRecords + 1

                 If Range("A" & j).Value Like "*" & PO Then

                    POSpent = POSpent + Range("D" & j).Value

                End If

                ActiveCell.Offset(1, 0).Select

            Next j

            Sheets("PO Management").Select

            If POSpent > 0 Then
                Range("U" & i).Value = POSpent
            End If

            ActiveCell.Offset(1, 0).Select
        Next i


    For k = 5 To nbPO + 4
        Range("T" & k).Select
        PO = ActiveCell.Value

        Sheets("InterDecSpending").Select

        Dim nbRecords2 As Integer

        Range("A2").Select
        nbRecords2 = Range(Selection, Selection.End(xlDown)).Rows.Count

        Dim POSpent2 As Integer

        POSpent2 = 0

        Dim l As Integer

            For l = 2 To nbRecords2 + 1

                 If Range("A" & l).Value Like "*" & PO Then

                    POSpent2 = POSpent2 + Range("D" & l).Value

                End If

                ActiveCell.Offset(1, 0).Select

            Next l

            Sheets("PO Management").Select

            If POSpent2 > 0 Then
                Range("U" & k).Value = POSpent2
            End If

            ActiveCell.Offset(1, 0).Select
        Next k

    For m = 5 To nbPO + 4
        Range("T" & m).Select
        PO = ActiveCell.Value

        Sheets("LaSalleSpending").Select

        Dim nbRecords3 As Integer

        Range("A2").Select
        nbRecords3 = Range(Selection, Selection.End(xlDown)).Rows.Count

        Dim POSpent3 As Integer

        POSpent3 = 0

        Dim n As Integer

            For n = 2 To nbRecords3 + 1

                 If Range("A" & n).Value Like "*" & PO Then

                    POSpent3 = POSpent3 + Range("D" & n).Value

                End If

                ActiveCell.Offset(1, 0).Select

            Next n

            Sheets("PO Management").Select

            If POSpent3 > 0 Then
                Range("U" & m).Value = POSpent3
            End If

            ActiveCell.Offset(1, 0).Select
        Next m

    For o = 5 To nbPO + 4
        Range("T" & o).Select
        PO = ActiveCell.Value

        Sheets("eLearningSpending").Select

        Dim nbRecords4 As Integer

        Range("A2").Select
        nbRecords4 = Range(Selection, Selection.End(xlDown)).Rows.Count

        Dim POSpent4 As Integer

        POSpent4 = 0

        Dim p As Integer

            For p = 2 To nbRecords4 + 1

                 If Range("A" & p).Value Like "*" & PO Then

                    POSpent4 = POSpent4 + Range("D" & p).Value

                End If

                ActiveCell.Offset(1, 0).Select

            Next p

            Sheets("PO Management").Select

            If POSpent4 > 0 Then
                Range("U" & o).Value = POSpent4
            End If

            ActiveCell.Offset(1, 0).Select
        Next o



End Sub
