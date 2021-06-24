VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "worksheetMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Worksheet As Worksheet
Attribute Worksheet.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set Worksheet = ActiveSheet
End Sub

Private Sub Class_Terminate()
    Cells.Interior.ColorIndex = 0
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.CountLarge > 1 Then Exit Sub
    
    Dim Saved_ws As Worksheet
    Dim Saved_Row As Integer
    Dim Saved_Column As Integer
    
    Set Saved_ws = Worksheets("worksheet_backup_hidden")
    
    Saved_Row = CInt(GetSetting("Crosshairs", "Startup", "Row"))
    Saved_Column = CInt(GetSetting("Crosshairs", "Startup", "Column"))
    
    
    If GetSetting("Crosshairs", "FilterSetting", "FilterMode") = "True" Then
        Call Restore_Visible_Formatting_Filtered(ActiveSheet, Saved_ws, Saved_Row, Saved_Column)
    Else
        Call Restore_Visible_Formatting(ActiveSheet, Saved_ws, Saved_Row, Saved_Column)
    End If
    
    With Target
        SaveSetting "Crosshairs", "Startup", "Row", .Row
        SaveSetting "Crosshairs", "Startup", "Column", .Column
            
        ' Highlight the entire row that contain the active cell
        ActiveSheet.Range(ActiveSheet.Cells(Target.Row, ActiveWindow.VisibleRange.Column), ActiveSheet.Cells(Target.Row, ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count)).Interior.ColorIndex = 4
        ActiveSheet.Range(ActiveSheet.Cells(ActiveWindow.VisibleRange.Row, Target.Column), ActiveSheet.Cells(ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count, Target.Column)).Interior.ColorIndex = 4
    End With
    
End Sub


Private Sub Restore_Visible_Formatting_Filtered(ws_active, ws_hidden, row_idx, col_idx)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim Active_Window_Last_Row As Integer
    Dim Active_Window_Last_Column As Integer
    
    Active_Window_Last_Row = ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count
    Active_Window_Last_Column = ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count
    
    Dim cell As Range
    
    For Each cell In ws_active.Range(ws_active.Cells(row_idx, ActiveWindow.VisibleRange.Column), ws_active.Cells(row_idx, Active_Window_Last_Column)).SpecialCells(xlCellTypeVisible)
        ws_hidden.Cells(cell.Row, cell.Column).Copy Destination:=cell
    Next cell
    
    For Each cell In ws_active.Range(ws_active.Cells(ActiveWindow.VisibleRange.Row, col_idx), ws_active.Cells(Active_Window_Last_Row, col_idx)).SpecialCells(xlCellTypeVisible)
        ws_hidden.Cells(cell.Row, cell.Column).Copy Destination:=cell
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub Restore_Visible_Formatting(ws_active, ws_hidden, row_idx, col_idx)
    Dim Active_Window_Last_Row As Integer
    Dim Active_Window_Last_Column As Integer
    
    Active_Window_Last_Row = ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count
    Active_Window_Last_Column = ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count
    
    ws_hidden.Range(ws_hidden.Cells(ActiveWindow.VisibleRange.Row, col_idx), ws_hidden.Cells(Active_Window_Last_Row, col_idx)).Copy Destination:=ws_active.Range(ws_active.Cells(ActiveWindow.VisibleRange.Row, col_idx), ws_active.Cells(Active_Window_Last_Row, col_idx))
    ws_hidden.Range(ws_hidden.Cells(row_idx, ActiveWindow.VisibleRange.Column), ws_hidden.Cells(row_idx, Active_Window_Last_Column)).Copy Destination:=ws_active.Range(ws_active.Cells(row_idx, ActiveWindow.VisibleRange.Column), ws_active.Cells(row_idx, Active_Window_Last_Column))
    
End Sub

