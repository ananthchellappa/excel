
Attribute VB_Name = "TableSerialNumberMacros"
Option Explicit

Sub FillInSerialNumbersFromTopMostCellInTable()
    Dim SelectedCell As Range
    Dim TableName As String, ColumnName As String
    Dim ActiveTable As ListObject
    Dim Rng As Range, ColumnHeader As Range
    Dim c As Long
    Dim dataArray() As Variant
    Dim i As Long
    Dim proceed As VbMsgBoxResult

    On Error GoTo NoTableSelected
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Identify the active cell and associated table
    Set SelectedCell = ActiveCell
    TableName = SelectedCell.ListObject.Name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    On Error GoTo 0

    ' Get the column header and the data body range
    Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)
    ColumnName = ColumnHeader.Value

    ' Check if column header is named '#'
    If ColumnName <> "#" Then
        ' Prompt user for confirmation
        proceed = MsgBox("The column header is not named '#'. Do you want to proceed?", _
                         vbYesNo + vbQuestion, "Confirm Action")
        If proceed = vbNo Then
            GoTo ExitSub
        End If
    End If

    ' Set the data range of the column
    Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange

    ' Load range into an array for faster updates
    dataArray = Rng.Value

    ' Fill the array with serial numbers
    For i = 1 To UBound(dataArray, 1)
        dataArray(i, 1) = i
    Next i

    ' Write the array back to the range in one operation
    Rng.Value = dataArray

ExitSub:
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

NoTableSelected:
    MsgBox "Please put cursor in a Table and run!", vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub FillInSerialNumbersFromLowestCellInTable()
    Dim SelectedCell As Range
    Dim TableName As String, ColumnName As String
    Dim ActiveTable As ListObject
    Dim Rng As Range, ColumnHeader As Range
    Dim c As Long
    Dim dataArray() As Variant
    Dim i As Long
    Dim proceed As VbMsgBoxResult

    On Error GoTo NoTableSelected
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Identify the active cell and associated table
    Set SelectedCell = ActiveCell
    TableName = SelectedCell.ListObject.Name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    On Error GoTo 0

    ' Get the column header and the data body range
    Set ColumnHeader = Intersect(ActiveCell.ListObject.HeaderRowRange, ActiveCell.EntireColumn)
    ColumnName = ColumnHeader.Value

    ' Check if column header is named '#'
    If ColumnName <> "#" Then
        ' Prompt user for confirmation
        proceed = MsgBox("The column header is not named '#'. Do you want to proceed?", _
                         vbYesNo + vbQuestion, "Confirm Action")
        If proceed = vbNo Then
            GoTo ExitSub
        End If
    End If

    ' Set the data range of the column
    Set Rng = ActiveTable.ListColumns(ColumnName).DataBodyRange
    c = Rng.Cells.Count

    ' Load range into an array for faster updates
    dataArray = Rng.Value

    ' Fill the array with reverse serial numbers
    For i = 1 To UBound(dataArray, 1)
        dataArray(i, 1) = c
        c = c - 1
    Next i

    ' Write the array back to the range in one operation
    Rng.Value = dataArray

ExitSub:
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

NoTableSelected:
    MsgBox "Please put cursor in a Table and run!", vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
