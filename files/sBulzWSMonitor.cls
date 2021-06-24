VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sBulzWSMonitor"
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
    
    Set Saved_ws = Worksheets("worksheet_backup_hidden")
    
    Saved_Row = CInt(GetSetting("Crosshairs", "Startup", "Row"))
    
    If Target.Row <> Saved_Row Then
     
        Call Restore_Visible_Formatting(ActiveSheet, Saved_ws, Saved_Row)
    
        With Target
            SaveSetting "Crosshairs", "Startup", "Row", .Row
            
            ' Highlight the entire row that contain the active cell
            ActiveSheet.Range(ActiveSheet.Cells(Target.Row, ActiveWindow.VisibleRange.Column), ActiveSheet.Cells(Target.Row, ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count)).Interior.ColorIndex = 4
        End With
    End If
    
End Sub


Private Sub Restore_Visible_Formatting(ws_active, ws_hidden, row_idx)
    Dim Active_Window_Last_Column As Integer
    
    Active_Window_Last_Column = ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count
    
    ws_hidden.Range(ws_hidden.Cells(row_idx, ActiveWindow.VisibleRange.Column), ws_hidden.Cells(row_idx, Active_Window_Last_Column)).Copy Destination:=ws_active.Range(ws_active.Cells(row_idx, ActiveWindow.VisibleRange.Column), ws_active.Cells(row_idx, Active_Window_Last_Column))
    
End Sub

