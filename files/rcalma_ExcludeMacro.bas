Attribute VB_Name = "Module1"
Option Explicit

Sub Tbl_Fil_DD_ON()

    On Error Resume Next
    ActiveCell.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveCell.AutoFilter Field:=1, Criteria1:="<>"
    ActiveCell.AutoFilter Field:=1
    If Err <> 0 Then MsgBox "Select any cell in table/data you want to filter and click the button" ' "AutoFilter Notification"
    On Error GoTo 0

End Sub

Sub Tbl_Fil_DD_OFF()

    On Error Resume Next
    ActiveCell.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveCell.AutoFilter Field:=1, Criteria1:="<>"
    ActiveCell.AutoFilter Field:=1
    Selection.AutoFilter
    If Err <> 0 Then MsgBox "Select any cell in table/data you want to filter and click the button" ' "AutoFilter Notification"
    On Error GoTo 0

End Sub

Sub clear_filter()
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    Tbl_Fil_DD_OFF
End Sub


Public Sub ExcludeFromFilter()
Attribute ExcludeFromFilter.VB_ProcData.VB_Invoke_Func = "E\n14"

    'On Error GoTo ErrorHandler
    
    '/// Test to see if the selector is currently in range of the table
    Dim TableIntersectionTest   As Range
    Set TableIntersectionTest = Intersect(ActiveSheet.ListObjects(ActiveCell.ListObject.Name).Range, Selection)
                                                                    'The Intersect function test if the selector |
                                                                                'intersects with the table range. If it
                                                                                'evaluates to nothing, it's not.
    
    If TableIntersectionTest Is Nothing Then
        MsgBox "Please place the cursor within the table range.", vbCritical
        Exit Sub
    End If
    '//////////
    Application.ScreenUpdating = False
    If Not ActiveSheet.FilterMode Then      ' 8/5/2020
        Tbl_Fil_DD_ON
    End If
    
    

    Dim CurrentFilter As Filter         'This is the iterator that will allow us to iterate through me.autofilter.filters
    
    Dim Criteria As Variant             'This is a local variable reflecting the .Criteria1 of the CurrentFilter
    
    Dim Criteria2 As Variant            'This is a local variable reflecting the .Criteria2 of the CurrentFilter.           |
                                        'See notes below. This will become important if there are only two filtered items.  |
                            
    'Dim IndexToRemove As Integer    '
    Dim ExclusionItem As Variant        'This is the string of the item we want removed
                                        'Ronald 08/16/20 Change variable from string to variant
    
    Dim CurrentColumnIndex As Integer   'This will help us track which column in the table we want to filter.   |
                                        'The Index reflects its location among the column headers.              |
    
    Dim OnlyTwoItems As Boolean         'This will test if we only have two items. Excel treats these filters differently.
    
    Dim NewCriteria As Variant          'This is a variant datatype that will store our new filter set. Could be a single |
                                        'value or an Array. See notes below.
    
    Dim TableHasFilters As Boolean      'This will test if the table has anything filtered. If not, this macro won't work.

    ExclusionItem = IIf(IsEmpty(Selection.Value), "", Application.WorksheetFunction.Text(Selection.Value, Selection.NumberFormat))  'Here we assigne the current selection value to the ExcelusionItem variable
                                                                                                   'Ronald 08/15/20 Added Cell formatting on ExclusionItem
                                                                                                   'Ronald 08/16/20 Added condition for blank entries
    
    
    '///    In this next bit of code, we need to find which column has been filtered. Even if           |
    '       there are dropdowns showing in the column headers, nothing may be filtered. To test         |
    '       if a filter has been deployed, we must itereate through all columns filters.                |
    '       In each case, we test if CurrentFilter.On, which evaluates to TRUE or FALSE. If it's        |
    '       true, we know we've hit the column we're interested in.                                     |
    
    'CurrentColumnIndex = ActiveSheet.ListObjects(ActiveCell.ListObject.Name).ListColumns(Selection.End(xlUp).Value).Index
    CurrentColumnIndex = 1 + ActiveCell.Column - ActiveSheet.ListObjects(ActiveCell.ListObject.Name).Range.Column
    
    Set CurrentFilter = ActiveSheet.AutoFilter.Filters(CurrentColumnIndex)
    'If CurrentFilter.On Then
    '    TableHasFilters = True
    '    If CurrentFilter.Count > 2 Then
    '        Criteria = CurrentFilter.Criteria1
    '    Else
    '        OnlyTwoItems = True
    '        Criteria = CurrentFilter.Criteria1
    '        Criteria2 = CurrentFilter.Criteria2
    '    End If
    'Else
        Criteria = getCriteria(ActiveSheet.ListObjects(ActiveCell.ListObject.Name).ListColumns(CurrentColumnIndex).Range.Address) 'Ronald 08/15/20 Create a new function to replace Application.Transpose
        Call QuickSort(Criteria, 1, UBound(Criteria))
        Call RemoveDuplicates(Criteria)
    'End If
    
    
    '////////////////// Table filtering code                                                            |
    '                                                                                                   |
    'Excel column filtering can be tricky. When there is only one item to be filtered,                  |
    'the filter object stores this data in its .Criteria1 fields. When there are two                    |
    'items to be filtered, these two items are stored in the .Criteria1 and .Criteria2                  |
    'respectively. However, when three or more items are part of the filter, the filtered               |
    'items are stored as an array in just .Criteria1.                                                   |
    '                                                                                                   |
    'In additiona to all of this, there is no way to turn one item on and off within a filter.          |
    'For instance, I can't use the code to just "turn off," as it were a specific item. Rather,         |
    'I have to clear out filter itself and then rebuild it without the item. In addition, to assign     |
    'a filter of three or more items, I have to build a string e.g. "1,2,4,5" of the filtered items     |
    'and feed it into an array and then assign to a totally new filter.                                 |
    '                                                                                                   |
    ' @i - this is a simple iterator. at its max, it will match the total count of filtered items       |
    ' @iNewCount - this is the new count. it should always reseolve to i - 1, if we're just removing    |
    '              one item                                                                             |
    ' @NewArrayString - NewArrayString stores what will be supplied next to the filter. It's a variant  |
    '                   which means it can take on either the form of a single variable or an           |
    '                   undimensionalized array at runtime.                                             |
    
    
    Dim i As Long
    Dim iNewCount As Long
    
    iNewCount = 1
    Dim NewArrayString
 
    'ActiveSheet.AutoFilter.ShowAllData           'Clearing the filter out -- commented out 9/18/19 - so that we keep the existing filtering
    If Not OnlyTwoItems Then            'Here we test if there are only two items. If so, the rules are         |
                                        'different. What follows is if there are more than two. |
                                        
        ReDim NewCriteria(1 To UBound(Criteria)) 'Ronald 08/15/20 change "UBound(Criteria) - 1" to "UBound(Criteria) to fixed issue #2"
        
        For i = LBound(Criteria) To UBound(Criteria)
            If Not Criteria(i) = "=" & ExclusionItem Then                'Filters are stored in Excel as "={item1},{item2}"
                NewCriteria(iNewCount) = Replace(Criteria(i), "=", "")  'As we assign the new items, we're remove that ='s
                Debug.Print "criteria list: " & NewCriteria(iNewCount) & " : " & Replace(Criteria(i), "=", "") & " : " & Criteria(i)
                iNewCount = iNewCount + 1
            End If
        Next
    Else  'Else condition is triggered if there are only two items
        
        'Here we test the two items .Criteria1 and .Criteria2 assigned respecitvely to Criteria and Criteria2.
        'If the excluded item matches Criteria, then we know Criteria2 is what we want to leave selected, and
        'vise versa
        
        'Added "=" on the line below. 'Ronald Calma
        If Criteria = "=" & ExclusionItem Then
            NewCriteria = Criteria2
        Else
            NewCriteria = Criteria
        End If
    End If

    'After all of this, we'll reapply the filters sans the excluded value. In this configuration,
    'we don't really care if NewCriteria is an array or a single value. We can assign it just
    'the same, and Excel will accept it.
    
    
    Selection.AutoFilter Field:=CurrentColumnIndex, _
        Criteria1:=NewCriteria, Operator:=xlFilterValues

    Application.ScreenUpdating = True

End Sub




Private Sub QuickSort(ByVal vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)
    
    While (tmpLow <= tmpHi)
       While (vArray(tmpLow) < pivot And tmpLow < inHi)
          tmpLow = tmpLow + 1
       Wend
    
       While (pivot < vArray(tmpHi) And tmpHi > inLow)
          tmpHi = tmpHi - 1
       Wend
    
       If (tmpLow <= tmpHi) Then
          tmpSwap = vArray(tmpLow)
          vArray(tmpLow) = vArray(tmpHi)
          vArray(tmpHi) = tmpSwap
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

Private Sub RemoveDuplicates(OldArray As Variant)

    Dim ArrayUpperBound As Long
    Dim NewUpperBound As Long
    
    Dim i As Long
    Dim UniqueValueCount As Long
    ReDim NewArrayOfUniques(1 To UBound(OldArray)) As Variant
    
    For i = 1 To UBound(OldArray)
        If i = 1 Then
            UniqueValueCount = 1
            NewArrayOfUniques(i) = OldArray(i)
        Else
            If OldArray(i) <> OldArray(i - 1) Then
                UniqueValueCount = UniqueValueCount + 1
                NewArrayOfUniques(UniqueValueCount) = OldArray(i)
            End If
        End If
    Next
    
    ReDim FinalArray(1 To UniqueValueCount)
    
    For i = 1 To UBound(FinalArray)
        FinalArray(i) = "=" & NewArrayOfUniques(i)
    
    Next
    
    OldArray = FinalArray
    
End Sub

'Ronald 08/15/20 Added this function to replace the Application.Tanspose function.
Function getCriteria(addr As String) As Variant
    Dim inp() As String
    Dim x As Long
    Dim cv As String
    Dim res() As Variant
    Dim st As Long
    Dim en As Long
    inp = Split(Replace(addr, ":", ""), "$") 'Split cell address by "$" to get the start and end rows
    st = Val(inp(2)) + 1 'start row handler. Remove the header row in the selection
    en = Val(inp(4)) 'start row handler.
    ReDim res(0)
    For x = st To en
        If Not ActiveSheet.Rows(Val(x) & ":" & Val(x)).Hidden Then 'Process only the rows that are not hidden.
        ReDim Preserve res(UBound(res) + 1)
          If IsEmpty(ActiveSheet.Range(inp(1) & x).Value) Then  'Ronald 08/16/20 Added condition for blank entries
            res(UBound(res)) = ""
          Else
            res(UBound(res)) = Application.WorksheetFunction.Text(ActiveSheet.Range(inp(1) & x).Value, ActiveSheet.Range(inp(1) & x).NumberFormat) 'get the formatted cell value.
          End If
          'Debug.Print res(UBound(res))
        End If
    Next x
    getCriteria = res
End Function

