' split string by list of words

Sub Main()
Dim ThisRng As Range
For iRow = 11 To 37405
    Set ThisRng = Range("C" & iRow)
    Call SplitContent(ThisRng)
Next
End Sub

Sub SplitContent(rng As Range)
Dim kwRng As Range
Dim arrSplitStrings() As String

Set kwRng = ThisWorkbook.Sheets("Sheet2").Range("A2:A17")

ThisStr = rng.Value
For Each oCell In kwRng
    ThisKw = oCell.Value
    ThisStr = Replace(ThisStr, ThisKw, "|")
Next
arrSplitStrings = Split(ThisStr, "|")
For i = 0 To UBound(arrSplitStrings)
rng.Offset(0, 4 + i).Value = arrSplitStrings(i)
Next


End Sub
