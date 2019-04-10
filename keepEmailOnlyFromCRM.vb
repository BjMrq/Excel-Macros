Sub A___keepEmailOnlyforFacebookAudiences()

        Columns("A:A").EntireColumn.Delete
        Columns("A:A").EntireColumn.Delete
        Columns("A:A").EntireColumn.Delete

    Dim nbColumns As Integer
    Range("A1").Select
    nbColumns = Range(Selection, Selection.End(xlToRight)).Columns.Count

    Range("A1").Select

    Dim i As Integer

    For i = 1 To nbColumns

        If ActiveCell.Value = "E-mail 1" Then
            ActiveCell.Value = "Email"
            ActiveCell.Offset(0, 1).Select
        Else
            ActiveCell.EntireColumn.Delete
        End If

    Next i

    Message = "How do you want to name your file?"
    Title = "detailListInput"

    nameOfImport = InputBox(Message, Title)


    ChDir "C:\Users\bmarquis\OneDrive - College LaSalle\Downloads"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\bmarquis\OneDrive - College LaSalle\Downloads\" & nameOfImport & ".csv" _
        , FileFormat:=xlCSVUTF8, CreateBackup:=False

End Sub
