VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Save_Filter_Form 
   Caption         =   "Save Filter"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "Save_Filter_Form.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Save_Filter_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Sub Save_Filter_Form_Show(Optional table_name As String, Optional filter_name As String, Optional notes As String, Optional criteria As String)
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If table_name <> "" Then
        tb_tablename.Text = table_name
        tb_filtername.Text = filter_name
        tb_notes.Text = notes
        tb_criteria.Text = criteria
    Else
        Dim Worksheet_Filters_Exists As Boolean
        Worksheet_Filters_Exists = Evaluate("ISREF('Saved_Filters'!A1)")
        
        ' If the saved filters worksheet doesn't exist in ActiveWorkbook then it means this is the first run and it has to add the necessary worksheets
        If Worksheet_Filters_Exists = False Then
            ActiveWorkbook.Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
            ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Name = "Saved_Filters"
            ActiveWorkbook.Worksheets("Saved_Filters").Visible = xlHidden
            
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 1) = "Bound Table"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 2) = "|"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 3) = "Filter Name"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 4) = "|"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 5) = "Notes"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 6) = "|"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 7) = "Criteria"
            ActiveWorkbook.Worksheets("Saved_Filters").Cells(1, 8) = "True Criterias"
            
            ActiveWorkbook.Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
            ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Name = "Remove_Duplicates"
            ActiveWorkbook.Worksheets("Remove_Duplicates").Visible = xlHidden
            
            
            ActiveWorkbook.Save
        End If
    
        Dim Table_Intersect As Variant
        Dim ActiveCell_List_Object As Variant
        Set ActiveCell_List_Object = ActiveCell.ListObject
        
        If ActiveCell_List_Object Is Nothing Then
            MsgBox "Please place the cursor in a table.", vbCritical
            
            Exit Sub
        End If
        
        Set Table_Intersect = Intersect(ws.ListObjects(ActiveCell_List_Object.Name).Range, Selection)
        
        tb_tablename.Text = Table_Intersect.ListObject.DisplayName
        
        Dim Full_Criteria_String As String
        Dim Full_True_Criteria_String As String
        
        Dim Header_Intersect_Range As Range
        Set Header_Intersect_Range = ws.ListObjects(ActiveCell.ListObject.Name).HeaderRowRange
        
        Dim Current_Filter As Filter
        
        Dim Criterias_Dictionary As Scripting.Dictionary
        Dim Current_Criteria As Variant
        Dim Row As Integer
        Dim Saved_Column As Variant
        Dim Missing_Item_Found As Boolean
        Dim Found_Needed_Criteria As String
       
        Full_Criteria_String = ""
        Full_True_Criteria_String = ""
        
        If ws.FilterMode = False Then
            MsgBox "Table does not have any filters.", vbCritical
            
            Exit Sub
        End If
        
        For i = 1 To ws.AutoFilter.Filters.Count
            Set CurrentFilter = ws.AutoFilter.Filters(i)
            
            If CurrentFilter.On = True Then
                
                If CurrentFilter.Count > 1 Then
                    Set Criterias_Dictionary = New Scripting.Dictionary
                    
                    If CurrentFilter.Count = 2 Then
                        Criterias_Dictionary.Add Replace(CStr(CurrentFilter.Criteria1), "=", ""), True
                        Criterias_Dictionary.Add Replace(CStr(CurrentFilter.Criteria2), "=", ""), True
                    Else
                        For Each Current_Criteria In CurrentFilter.Criteria1
                            Criterias_Dictionary.Add Replace(CStr(Current_Criteria), "=", ""), True
                        Next Current_Criteria
                    End If
                    
                    ' Current Column: Header_Intersect_Range.Column + i - 1
                    
                    Saved_Column = ws.Range(ws.Cells(Header_Intersect_Range.Row, Header_Intersect_Range.Column + i - 1), ws.Cells(Header_Intersect_Range.Row + Table_Intersect.ListObject.DataBodyRange.Rows.Count, Header_Intersect_Range.Column + i - 1))
                    
                    ActiveWorkbook.Worksheets("Remove_Duplicates").Range("A1").Resize(UBound(Saved_Column, 1), UBound(Saved_Column, 2)).Value = Saved_Column
                
                    
                    ' Remove the duplicates in the current column in order to search for the missing items from the filter
                    
                    ActiveWorkbook.Worksheets("Remove_Duplicates").Range(ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(1, 1), ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(ActiveWorkbook.Worksheets("Remove_Duplicates").UsedRange.Rows.Count, 1)).RemoveDuplicates Columns:=1, Header:=xlYes
                    
                    ' Look for the missing items in order to see what has been subtracted from the filter
                    
                    Found_Needed_Criteria = ""
                    
                    Row = 2
                    Do While Row <= ActiveWorkbook.Worksheets("Remove_Duplicates").UsedRange.Rows.Count
                        If ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(Row, 1) <> "" Then
                            If Criterias_Dictionary.Exists(CStr(ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(Row, 1))) = False Then
                                Found_Needed_Criteria = Found_Needed_Criteria + CStr(ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(Row, 1)) + "|"
                            End If
                        End If
                        
                        Row = Row + 1
                    Loop
                    
                    Found_Needed_Criteria = Left(Found_Needed_Criteria, Len(Found_Needed_Criteria) - 1)
                    
                    Full_Criteria_String = Full_Criteria_String + CStr(ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(1, 1)) + "!=" + Found_Needed_Criteria + ";"
                    Full_True_Criteria_String = Full_True_Criteria_String + CStr(ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(1, 1)) + ":" + Join(CurrentFilter.Criteria1, "|") + ";"
                    
                    ActiveWorkbook.Worksheets("Remove_Duplicates").Columns(1).EntireColumn.Delete
                Else
                    If CStr(CurrentFilter.Criteria1) = "" Then
                        Full_Criteria_String = Full_Criteria_String + CStr(ws.Cells(Header_Intersect_Range.Row, Header_Intersect_Range.Column + i - 1)) + "=(Blanks)" + ";"
                    ElseIf CStr(CurrentFilter.Criteria1) = "=" Then
                        Full_Criteria_String = Full_Criteria_String + CStr(ws.Cells(Header_Intersect_Range.Row, Header_Intersect_Range.Column + i - 1)) + "=(Blanks)" + ";"
                    ElseIf CStr(CurrentFilter.Criteria1) = "<>" Then
                        Full_Criteria_String = Full_Criteria_String + CStr(ws.Cells(Header_Intersect_Range.Row, Header_Intersect_Range.Column + i - 1)) + "!=(Blanks)" + ";"
                    Else
                        Full_Criteria_String = Full_Criteria_String + CStr(ws.Cells(Header_Intersect_Range.Row, Header_Intersect_Range.Column + i - 1)) + CStr(CurrentFilter.Criteria1) + ";"
                    End If
                    
                    Full_True_Criteria_String = Full_True_Criteria_String + CStr(ws.Cells(Header_Intersect_Range.Row, Header_Intersect_Range.Column + i - 1)) + ":" + CurrentFilter.Criteria1 + ";"
                End If
                
            End If
            
        Next i
        
        tb_criteria.Text = Left(Full_Criteria_String, Len(Full_Criteria_String) - 1)
        tb_true_criteria.Text = Left(Full_True_Criteria_String, Len(Full_True_Criteria_String) - 1)
    End If
    
    
    Save_Filter_Form.Show
    
    

End Sub


Private Sub UserForm_Initialize()

    tb_criteria.BackColor = RGB(255, 165, 0)

End Sub

Private Sub cb_run_save_Click()

    Dim Counter_Rows_Saved_Filters As Integer
    Counter_Rows_Saved_Filters = 2
    
    Dim Entry_Found_Already As Boolean
    Dim Entry_Found_Row As Integer
    
    Entry_Found_Already = False
    
    Do While ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 1) <> ""
        If ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 1) = tb_tablename.Text And ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 3) = tb_filtername.Text Then
            Entry_Found_Already = True
            
            Entry_Found_Row = Counter_Rows_Saved_Filters
        End If
        
        Counter_Rows_Saved_Filters = Counter_Rows_Saved_Filters + 1
    Loop
    
    
    If Entry_Found_Already = False Then
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 1) = tb_tablename.Text
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 2) = "|"
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 3) = tb_filtername.Text
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 4) = "|"
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 5) = tb_notes.Text
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 6) = "|"
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 7) = tb_criteria.Text
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved_Filters, 8) = tb_true_criteria.Text
    Else
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Entry_Found_Row, 5) = tb_notes.Text
        ActiveWorkbook.Worksheets("Saved_Filters").Cells(Entry_Found_Row, 7) = tb_criteria.Text
    End If
    
    
    ActiveWorkbook.Worksheets("Saved_Filters").UsedRange.Columns.AutoFit
    
    ActiveWorkbook.Save
    
    Me.Hide

End Sub
