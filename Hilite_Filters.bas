Attribute VB_Name = "Hilite_Filters"
Option Explicit
Sub FilterOdd()
    
    Dim CurTable    As ListObject
    Dim Color       As String
    Dim ListRow     As Variant
    Dim CurCell     As Range
    Dim FilterArray()    As String        'for filter
    Dim Firstcell   As Range
    Dim arraycount  As Integer
    Dim ColCount    As Integer
    Dim RowCount    As Integer
    Dim Row         As Integer
    Dim Firstcolor  As Variant
    Dim Col         As Integer
    Dim K           As Variant
    Dim Hidden      As Boolean
    
    On Error GoTo Err
    
    'Check if Activecell is within table
    arraycount = 0
    If ActiveCell.ListObject.Name <> "" Then
        Set CurTable = ActiveCell.ListObject        'Save table
        ColCount = CurTable.ListColumns.count        'number of columns in table
        RowCount = CurTable.ListRows.count        ' number of rows in table
        
        'Loop through all rows
        For Row = 1 To RowCount
            Set ListRow = CurTable.ListRows(Row)
            'loop through all columns
            Set Firstcell = ListRow.Range.Cells(1, 1)
            
            Firstcolor = Firstcell.Interior.Color
            For Col = 1 To ColCount
                
                Set CurCell = ListRow.Range.Cells(1, Col)
                K = CurCell.Interior.Color
                Hidden = ListRow.Range.EntireRow.Hidden        'Check if filtered
                'Check If Hidden & Color is different
                If K <> Firstcolor And Not Hidden Then
                    ReDim Preserve FilterArray(arraycount)
                    FilterArray(arraycount) = Firstcell.Value
                    arraycount = arraycount + 1
                End If
            Next Col
            
        Next Row
        'update filter
        Application.ScreenUpdating = False
        If arraycount <> 0 Then
            CurTable.Range.AutoFilter field:=1, Criteria1:= _
                                      FilterArray, Operator:=xlFilterValues
        Else
            CurTable.Range.AutoFilter field:=1, Criteria1:= _
                                      "", Operator:=xlFilterValues
            
        End If
        Application.ScreenUpdating = True
        
    End If
    Exit Sub
Err:
    MsgBox "This Function requires Active cell within Table", , " Error fromOddCellFilter"
    Application.ScreenUpdating = True
End Sub

Sub FilterOddSelectedColumn()
    
    Dim CurTable    As ListObject
    Dim Color       As String
    Dim ListRow      As Variant
    Dim CurCell     As Range
    Dim SelectedCol As Integer
    Dim FilterArray()    As String        'for filter
    Dim Firstcell   As Range
    Dim arraycount  As Integer
    Dim ColCount    As Integer
    Dim RowCount    As Integer
    Dim Row         As Integer
    Dim Firstcolor  As Variant
    Dim Col         As Integer
    Dim K           As Variant
    Dim Hidden      As Boolean
    
    On Error GoTo Err
    'Check if Active is within table
    
    If ActiveCell.ListObject.Name <> "" Then
        Set CurTable = ActiveCell.ListObject
        
        ColCount = CurTable.ListColumns.count
        RowCount = CurTable.ListRows.count
        SelectedCol = ActiveCell.Column - CurTable.HeaderRowRange.Column + 1
        
        'Loop through all rows
        For Row = 1 To RowCount
            Set ListRow = CurTable.ListRows(Row)
            
            Set Firstcell = ListRow.Range.Cells(1, 1)
            
            Firstcolor = Firstcell.Interior.Color
            
            Set CurCell = ListRow.Range.Cells(1, SelectedCol)
            K = CurCell.Interior.Color
            Hidden = ListRow.Range.EntireRow.Hidden        'Check if filtered
            
            'Check If Hidden & Color is different
            If K <> Firstcolor And Not Hidden Then
                ReDim Preserve FilterArray(arraycount)
                FilterArray(arraycount) = CurCell.Value
                arraycount = arraycount + 1
            End If
            
        Next Row
        'update filter
        Application.ScreenUpdating = False
        If arraycount <> 0 Then
            CurTable.Range.AutoFilter field:=SelectedCol, Criteria1:= _
                                      FilterArray, Operator:=xlFilterValues
        Else
            CurTable.Range.AutoFilter field:=SelectedCol, Criteria1:= _
                                      "", Operator:=xlFilterValues
            
        End If
        Application.ScreenUpdating = True
        
    End If
    
    Exit Sub
Err:
    MsgBox "This Function requires Active cell within Table", , " Error from FilterOddSelectedColumn"
    Application.ScreenUpdating = True
End Sub



