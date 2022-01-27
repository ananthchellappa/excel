Sub DeleteNotVisibleRowsFromTable()
Dim SelectedCell As Range
Dim TableName As String, ColumnName As String
Dim ActiveTable As ListObject
Dim Rng As Range, cel As Range, ColumnHeader As Range
Dim b1 As Long, b2 As Long, i As Long, c As Long, col As Long, vRows() As Variant, b As Long

b = 0
Set SelectedCell = ActiveCell

On Error GoTo NoTableSelected
TableName = SelectedCell.ListObject.Name
Set ActiveTable = ActiveSheet.ListObjects(TableName)
On Error GoTo 0

If MsgBox("Will delete hidden rows. Proceed?", vbYesNo) = vbYes Then

Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)
b1 = ColumnHeader.Offset(1).Row
col = ColumnHeader.Column
ColumnName = ColumnHeader.Value
c = ActiveTable.ListColumns(ColumnName).DataBodyRange.Cells.count
b2 = b1 + c - 1

Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange

For i = b2 To b1 Step -1
If ActiveSheet.Cells(i, col).EntireRow.Hidden = True Then
ReDim Preserve vRows(b)
vRows(b) = ActiveSheet.Cells(i, col).Row - (b1 - 1)
'Debug.Print vRows(b)
b = b + 1
End If
Next i

If b > 0 Then
ActiveSheet.ListObjects(TableName).AutoFilter.ShowAllData
For i = LBound(vRows) To UBound(vRows)
ActiveSheet.ListObjects(TableName).ListRows(vRows(i)).Delete
Next i
End If

End If

Exit Sub
NoTableSelected:
MsgBox "Please put cursor in a Table and run!", vbCritical

End Sub


