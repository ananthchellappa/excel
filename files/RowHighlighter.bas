Attribute VB_Name = "RowHighlighter"
Option Explicit

Private Monitor As sBulzWSMonitor

Function checkHighlight() As Boolean
    Dim result As Boolean
    Dim c As Variant
    result = True
    For Each c In ActiveSheet.UsedRange
        If c.Interior.Pattern <> xlNone Then
            result = False
            Exit For
        End If
    Next c
    checkHighlight = result
    If result = False Then
        If MsgBox("Custom Highlighting detected. Proceed?", vbYesNo) = vbYes Then
            checkHighlight = True
        End If
    End If
End Function

Sub StartRowHighlight()
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim Saved_Info_ws As Worksheet
    
    ActiveSheet.Copy After:=ActiveSheet
    
    Set Saved_Info_ws = ActiveWorkbook.Worksheets(ws.Index + 1)
    
    
    Saved_Info_ws.Visible = xlSheetHidden
    Saved_Info_ws.Name = "worksheet_backup_hidden"
    
    ws.Activate
    
    If checkHighlight Then
        Set Monitor = New sBulzWSMonitor
        With ActiveCell
            SaveSetting "Crosshairs", "Startup", "Row", .Row
            
            ' Highlight the entire row that contain the active cell
            ActiveSheet.Range(ActiveSheet.Cells(ActiveCell.Row, ActiveWindow.VisibleRange.Column), ActiveSheet.Cells(ActiveCell.Row, ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count)).Interior.ColorIndex = 4
        End With
    End If
    

End Sub

Sub StopRowHighlight()

    Set Monitor = Nothing
    
    Dim Selected_Range_Row As Integer
    Dim Selected_Range_Column As Integer
    
    Selected_Range_Row = ActiveCell.Row
    Selected_Range_Column = ActiveCell.Column
    
    Dim Saved_ws As Worksheet
    Set Saved_ws = Worksheets("worksheet_backup_hidden")
    
    Application.DisplayAlerts = False
    
    Dim Worksheet_Name As String
    
    Worksheet_Name = ActiveSheet.Name
    
    Saved_ws.Visible = xlSheetVisible
    
    ActiveSheet.Delete
    
    Saved_ws.Name = Worksheet_Name
    
    DeleteSetting "Crosshairs", "Startup"
    
    Saved_ws.Cells(Selected_Range_Row, Selected_Range_Column).Select
    
    Application.DisplayAlerts = True
    
End Sub
