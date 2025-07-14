Sub RemoveFirstWord()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim col As String

    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Set the column to process (e.g., Column A)
    col = "A"
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Set rng = ws.Range(col & "1:" & col & lastRow)

    For Each cell In rng
        If InStr(cell.Value, " ") > 0 Then
            cell.Value = Mid(cell.Value, InStr(cell.Value, " ") + 1)
        End If
    Next cell
End Sub
