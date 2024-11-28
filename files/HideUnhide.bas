Attribute VB_Name = "HideUnhide"
Option Explicit

' Main subroutine
Sub HideSelectedColumnsFromTable()
    Dim ActiveTable As ListObject
    Dim ColumnHeader As Range
    Dim SelectedColumns As Collection
    Dim i As Integer
    Dim HeaderName As Variant
    Dim ActualColumnName As String ' Extracted column name from the ListBox entry

    On Error GoTo NotInTable
    ' Check if active cell is in a table
    Set ActiveTable = ActiveCell.ListObject
    On Error GoTo 0

    ' Show the UserForm
    Dim frm As New HideColumnsDialog
    frm.Show vbModal

    ' Check if OK was pressed; exit if not
    If Not frm.OKPressed Then
        Unload frm ' Explicitly unload the form
        Exit Sub
    End If

    ' Collect selected columns from the ListBox
    Set SelectedColumns = New Collection
    For i = 0 To frm.lstColumnNames.ListCount - 1
        If frm.lstColumnNames.Selected(i) Then
            SelectedColumns.Add frm.lstColumnNames.List(i)
        End If
    Next i

    ' If no columns were selected, exit silently
    If SelectedColumns.count = 0 Then
        Unload frm ' Explicitly unload the form
        Exit Sub
    End If

    ' Hide the selected columns
    For Each HeaderName In SelectedColumns
        ' Extract the actual column name by splitting the ListBox entry
        ActualColumnName = Trim(Split(HeaderName, " - ")(1))
        ActiveTable.ListColumns(ActualColumnName).Range.EntireColumn.Hidden = True
    Next HeaderName

    ' Display success message only if columns were hidden
    MsgBox "Selected columns have been hidden.", vbInformation, "Done"
    Unload frm ' Explicitly unload the form
    Exit Sub

NotInTable:
    MsgBox "Please place the cursor in an Excel table and run the subroutine.", vbCritical, "Error"
End Sub



Sub UnhideSelectedColumnsFromTable()
    Dim ActiveTable As ListObject
    Dim ColumnHeader As Range
    Dim HiddenColumnCount As Integer
    Dim SelectedColumns As Collection
    Dim i As Integer
    Dim HeaderName As Variant
    Dim ActualColumnName As String

    On Error GoTo NotInTable
    ' Check if active cell is in a table
    Set ActiveTable = ActiveCell.ListObject
    On Error GoTo 0

    ' Count hidden columns
    HiddenColumnCount = 0
    For Each ColumnHeader In ActiveTable.HeaderRowRange
        If ColumnHeader.EntireColumn.Hidden Then
            HiddenColumnCount = HiddenColumnCount + 1
        End If
    Next ColumnHeader

    ' If no hidden columns, exit silently
    If HiddenColumnCount = 0 Then
        MsgBox "No hidden columns are available to unhide.", vbInformation, "No Hidden Columns"
        Exit Sub
    End If

    ' Show the UserForm
    Dim frm As New UnhideColumnsDialog
    frm.Show vbModal

    ' Check if OK was pressed; exit if not
    If Not frm.OKPressed Then
        Unload frm ' Explicitly unload the form
        Exit Sub
    End If

    ' Collect selected columns from the ListBox
    Set SelectedColumns = New Collection
    For i = 0 To frm.lstHiddenColumns.ListCount - 1
        If frm.lstHiddenColumns.Selected(i) Then
            SelectedColumns.Add frm.lstHiddenColumns.List(i)
        End If
    Next i

    ' If no columns were selected, exit silently
    If SelectedColumns.count = 0 Then
        Unload frm ' Explicitly unload the form
        Exit Sub
    End If

    ' Unhide the selected columns
    For Each HeaderName In SelectedColumns
        ' Extract the actual column name by splitting the ListBox entry
        ActualColumnName = Trim(Split(HeaderName, " - ")(1))
        ActiveTable.ListColumns(ActualColumnName).Range.EntireColumn.Hidden = False
    Next HeaderName

    ' Display success message only if columns were unhidden
    MsgBox "Selected columns have been unhidden.", vbInformation, "Done"
    Unload frm ' Explicitly unload the form
    Exit Sub

NotInTable:
    MsgBox "Please place the cursor in an Excel table and run the subroutine.", vbCritical, "Error"
End Sub

