Attribute VB_Name = "Vineeth"
' code by Sai Vineeth Tallapudi
' Address the need for a "contains" version of the Filter By Selected cell and Exclude (filter out) by selected cell

Sub Filter_Include()
    Call Text_FilterHelper  'ALT-SHIFT-F
End Sub
Sub Filter_Include_CaseSensitive()
    Call Text_FilterHelper(False, True) 'ALT-SHIFT-T
End Sub
Sub Filter_Exclude()
    Call Text_FilterHelper(True)    'ALT-SHIFT-E
End Sub
Sub Filter_Exclude_CaseSensitive()
    Call Text_FilterHelper(True, True)  'ALT-SHIFT-R
End Sub

'Text_FilterHelper()            to filter based on selected text          - Not CaseSensitive
'Text_FilterHelper(False,True)  to filter based on selected text          - Case Sensitive
'Text_FilterHelper(True)        to filter exclude based on selected text  - Not CaseSensitive
'Text_FilterHelper(True,True)   to filter exclude based on selected text  - Case Sensitive
'

Sub Text_FilterHelper(Optional Exclude As Boolean, Optional isCaseSensitive As Boolean)
    
    Dim CurTable    As ListObject
    Dim filtercriteria() As Variant
    Dim CurColNum   As Integer
    Dim CurCol      As ListColumn
    Dim SNoCol      As ListColumn
    Dim Count       As Integer
    Dim lookupstring As String
    Dim IsMatch     As Boolean
    Dim Values      As Variant
    Dim SNoValues   As Variant
    Dim TCell       As Range
    Dim Value       As Variant
    Dim Filter      As Variant
    Dim RowIndex    As Integer
    
    
    Count = 0
    ' Get Current Table
    Set CurTable = ActiveSheet.ListObjects(ActiveCell.ListObject.Name)
    ' Get ColNum
    CurColNum = 1 + ActiveCell.Column - CurTable.Range.Column
    ' Get Table Colunm
    Set CurCol = CurTable.ListColumns.Item(CurColNum)
    ' Get Selected cell Value
    lookupstring = ActiveCell.Value
    
    ' Get Values of selected column in Current Table
    Set Values = CurCol.Range.Cells
    Set SNoCol = CurTable.ListColumns.Item(1)
    SNoValues = SNoCol.Range.Value
    
    ' If Not Case Sensitive then change text to upper
    If Not isCaseSensitive Then
        lookupstring = UCase(lookupstring)
    End If
    RowIndex = 1
    For Each TCell In Values
        Value = TCell.Value
        ' Since first value is Heading , we ignore when count =0
        If Count > 0 Then
            
            ' If Not Case Sensitive then change text to upper
            If Not isCaseSensitive Then
                Value = UCase(Value)
            End If
            
            ' Check Whether selected string exists given string
            IsMatch = InStr(Value, lookupstring) > 0
            
            ' If filter Type is include and matched
            ' then we add values to filtercriteria
            If Not Exclude And IsMatch Then
                ReDim Preserve filtercriteria(Count)
                filtercriteria(Count - 1) = CStr(SNoValues(RowIndex, 1))
                Count = Count + 1
                ' if Filter type is exclude and not matched and cell is not hidden
                ' then we add values to filtercriteria
            ElseIf Exclude And Not IsMatch And Not TCell.EntireRow.Hidden Then
                ReDim Preserve filtercriteria(Count)
                filtercriteria(Count - 1) = CStr(SNoValues(RowIndex, 1))
                Count = Count + 1
            End If
        Else
            Count = Count + 1
        End If
    RowIndex = RowIndex + 1
    Next
    ' Apply filter based on filtrer critertia
    ' Filter is applied on Selected column in Table
    CurTable.Range.AutoFilter Field:=1, Criteria1:= _
    filtercriteria, Operator:=xlFilterValues
End Sub







