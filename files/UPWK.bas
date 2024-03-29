Sub ActiveTableFillInBlanks()
Dim SelectedCell As Range
Dim TableName As String, ColumnName As String
Dim ActiveTable As ListObject
Dim Rng As Range, cel As Range, ColumnHeader As Range

Set SelectedCell = ActiveCell

On Error GoTo NoTableSelected
TableName = SelectedCell.ListObject.Name
Set ActiveTable = ActiveSheet.ListObjects(TableName)
On Error GoTo 0

Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)

ColumnName = ColumnHeader.Value
Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange

For Each cel In Rng
If Trim(cel.Value) = "" And cel.Address <> ColumnHeader.Offset(1).Address Then
cel.Value = cel.Offset(-1).Value
End If
Next cel

Exit Sub
NoTableSelected:
MsgBox "Please put cursor in a Table and run!", vbCritical

End Sub

Sub FillInSerialNumbersFromTopMostCellInTable()
Dim SelectedCell As Range
Dim TableName As String, ColumnName As String
Dim ActiveTable As ListObject
Dim Rng As Range, cel As Range, ColumnHeader As Range
Dim c As Long

Set SelectedCell = ActiveCell
c = 1

On Error GoTo NoTableSelected
TableName = SelectedCell.ListObject.Name
Set ActiveTable = ActiveSheet.ListObjects(TableName)
On Error GoTo 0

Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)

ColumnName = ColumnHeader.Value
Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange

For Each cel In Rng
cel.Value = c
c = c + 1
Next cel

Exit Sub
NoTableSelected:
MsgBox "Please put cursor in a Table and run!", vbCritical

End Sub


Sub FillInSerialNumbersFromLowestCellInTable()
Dim SelectedCell As Range
Dim TableName As String, ColumnName As String
Dim ActiveTable As ListObject
Dim Rng As Range, cel As Range, ColumnHeader As Range
Dim c As Long

Set SelectedCell = ActiveCell


On Error GoTo NoTableSelected
TableName = SelectedCell.ListObject.Name
Set ActiveTable = ActiveSheet.ListObjects(TableName)
On Error GoTo 0

Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)

ColumnName = ColumnHeader.Value
Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange
c = ActiveTable.ListColumns(ColumnName).DataBodyRange.Cells.count

For Each cel In Rng
cel.Value = c
c = c - 1
Next cel

Exit Sub
NoTableSelected:
MsgBox "Please put cursor in a Table and run!", vbCritical

End Sub

