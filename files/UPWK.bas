Attribute VB_Name = "UPWK"
Sub ActiveTableFillInBlanks()

Dim SelectedCell As Range
Dim TableName As String, ColumnName As String
Dim ActiveTable As ListObject
Dim Rng As Range, cel As Range, i As Long, ColumnHeader As Range

Set SelectedCell = ActiveCell
i = 1

On Error GoTo NoTableSelected
TableName = SelectedCell.ListObject.Name
Set ActiveTable = ActiveSheet.ListObjects(TableName)
On Error GoTo 0

Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)

ColumnName = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn).Value
Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange
Debug.Print Rng.Address

For Each cel In Rng
    If Trim(cel.Value) = "" And cel.Address <> ColumnHeader.Address Then
        cel.Value = cel.Offset(-1).Value
    End If
Next cel

Exit Sub
NoTableSelected:
MsgBox "Please put cursor in a Table and run!", vbCritical

End Sub
