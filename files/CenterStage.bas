Option Explicit
Sub CenterStage()
Dim rngFrozen As Range
Dim arrHCols() As String
Dim currentMode As String
Dim c As Long, hCol As Long, sCol As Long, eCol As Long
Dim keepHidden As Boolean
Dim doToggle As Boolean
  Set rngFrozen = GetFreezePaneRange
  currentMode = GetMode
  If rngFrozen.Column = 1 Or (currentMode = "Original" And Selection.Column <= rngFrozen.Column) Then Exit Sub
  
  'Functions GetHiddenColumns updates hidden columns variable if mode is original
  arrHCols = Split(GetHiddenColumns, ",")
  With ActiveSheet
    sCol = rngFrozen.Column
    
    'if mode is original and at least one cell in selection is not empty
    If currentMode = "Original" And Application.WorksheetFunction.CountA(Selection) > 0 Then
      eCol = Selection.Column - 1
      If sCol < eCol Then
        For c = sCol To eCol
          If .Cells(1, c).EntireColumn.Hidden = False Then
            .Cells(1, c).EntireColumn.Hidden = True
            doToggle = True
          End If
        Next
      End If
      
    Else
      If currentMode = "Presentation" Then
        eCol = .UsedRange.Columns.count
        If sCol < eCol Then
          For c = sCol To eCol
            If .Cells(1, c).EntireColumn.Hidden = False Then GoTo NextColumn
            keepHidden = False
            For hCol = LBound(arrHCols) To UBound(arrHCols)
              If c = CLng(arrHCols(hCol)) Then
                keepHidden = True
                Exit For
              End If
            Next
            If Not keepHidden Then
              .Cells(1, c).EntireColumn.Hidden = False
              doToggle = True
            End If
NextColumn:
          Next
        End If
        
      End If
    End If
  End With
  If doToggle Then ToggleMode
End Sub
Function GetHiddenColumns() As String
  Const sHCols As String = "HiddenCols"
  Dim cstmDocProp As DocumentProperty
  Dim strCols As String
  Dim sCol As Long, eCol As Long, c As Long
  Dim currentMode As String
  
  With ActiveSheet
    sCol = GetFreezePaneRange.Column
    eCol = .UsedRange.Columns.count
    
    'if mode is original then get all existing hidden columns
    currentMode = GetMode
    If currentMode = "Original" And sCol < eCol Then
      For c = sCol To eCol
        If .Cells(1, c).EntireColumn.Hidden = True Then
          If strCols = "" Then
            strCols = CStr(c)
          Else
            strCols = strCols & "," & CStr(c)
          End If
        End If
      Next
    End If
    
  End With
  
  On Error Resume Next
  Set cstmDocProp = ActiveWorkbook.CustomDocumentProperties(sHCols)
  
  
  'If the property doesn't exist, create it and set the initial value
  If Err.Number > 0 Then
    ActiveWorkbook.CustomDocumentProperties.Add _
    Name:=sHCols, _
    LinkToContent:=False, _
    Type:=msoPropertyTypeString, _
    Value:=strCols
    Set cstmDocProp = ActiveWorkbook.CustomDocumentProperties(sHCols)
  End If
  On Error GoTo 0
  
  'If sheet is in original mode then get currently hidden columns else get the hidden columns
  'when sheet was in original mode previously
  If currentMode = "Original" Then
    cstmDocProp.Value = strCols
    GetHiddenColumns = strCols
  Else
    GetHiddenColumns = cstmDocProp.Value
  End If
  
  Set cstmDocProp = Nothing
End Function
Function GetMode() As String
  Const sMode As String = "SheetMode"
  Dim cstmDocProp As DocumentProperty
  
  'If the name doesn't exist, create it and set the initial value to Original
  On Error Resume Next
  Set cstmDocProp = ActiveWorkbook.CustomDocumentProperties(sMode)
  If Err.Number > 0 Then
    ActiveWorkbook.CustomDocumentProperties.Add _
    Name:=sMode, _
    LinkToContent:=False, _
    Type:=msoPropertyTypeString, _
    Value:="Original"
    Set cstmDocProp = ActiveWorkbook.CustomDocumentProperties(sMode)
  End If
  
  GetMode = cstmDocProp.Value
  Set cstmDocProp = Nothing
End Function
Sub ToggleMode()
  Const sMode As String = "SheetMode"
     
  'If the name doesn't exist, we create it and set the initial value to 1
  On Error Resume Next
  Dim cstmDocProp As DocumentProperty
  Set cstmDocProp = ActiveWorkbook.CustomDocumentProperties(sMode)
  If Err.Number > 0 Then
    ActiveWorkbook.CustomDocumentProperties.Add _
    Name:=sMode, _
    LinkToContent:=False, _
    Type:=msoPropertyTypeString, _
    Value:="Original"
  Else
    On Error GoTo 0
    'if mode property exists, toggle the value
    Dim sModeVal As String
    sModeVal = ActiveWorkbook.CustomDocumentProperties(sMode).Value

    '   Toggle the mode value
    If sModeVal = "Original" Then
      ActiveWorkbook.CustomDocumentProperties(sMode).Value = "Presentation"
    Else
      ActiveWorkbook.CustomDocumentProperties(sMode).Value = "Original"
    End If
         
  End If
  Set cstmDocProp = Nothing
End Sub

Function GetFreezePaneRange() As Range
  Dim Rw As Long, Col As Long

  With ActiveWindow
    If .SplitColumn <> 0 Then
      Rw = .SplitRow + 1
      Col = .SplitColumn + 1
      'get range of top left cell below the freeze pane
      Set GetFreezePaneRange = ActiveSheet.Cells(Rw, Col)
    Else
      Set GetFreezePaneRange = ActiveSheet.Range("A1")
    End If
  End With
  
End Function



