Sub cleanDataErrorFacebookLeadGen()
'
' cleanDataErrorFacebookLeadGen Macro
'

'
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1)), _
        TrailingMinusNumbers:=True

    Columns("B:P").Select
    Selection.Delete Shift:=xlToLeft

    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft

    ActiveWorkbook.SaveAs Filename:= _
          "C:\Users\bmarquis\OneDrive - College LaSalle\Downloads\ErrorsReadyToImportAgain.csv", FileFormat _
          :=xlCSVUTF8, CreateBackup:=False
End Sub
