VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} View_Filters_Form 
   Caption         =   "View Saved Filters"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   OleObjectBlob   =   "View_Filters_Form.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "View_Filters_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Sub View_Filters_Show_Form()

    Dim Worksheet_Filters_Exists As Boolean
    Worksheet_Filters_Exists = Evaluate("ISREF('Saved_Filters'!A1)")
    
    If Worksheet_Filters_Exists = False Then
        MsgBox ("There are no saved filters for this file."), vbCritical
        
        Exit Sub
    Else
        Dim ws_saved_fil As Worksheet
        Set ws_saved_fil = ActiveWorkbook.Worksheets("Saved_Filters")
        
        If ws_saved_fil.Cells(2, 1) = "" Then
            MsgBox ("There are no saved filters for this file."), vbCritical
        
            Exit Sub
        End If
        
        Dim Table_Intersect As Range
        Dim ActiveCell_List_Object As Variant
        Set ActiveCell_List_Object = ActiveCell.ListObject
        
        cb_tablename.AddItem "All"
        
        Dim Counter_Rows_Saved As Integer
        Counter_Rows_Saved = 2
        
        Dim Saved_Column As Variant
        
        Saved_Column = ws_saved_fil.Range(ws_saved_fil.Cells(2, 1), ws_saved_fil.Cells(ws_saved_fil.UsedRange.Rows.Count, 1))
        
        Dim Aux_Saved As Variant
        If IsArray(Saved_Column) = False Then
            Aux_Saved = Saved_Column
            
            ReDim Saved_Column(1 To 1, 1 To 1)
            Saved_Column(1, 1) = Aux_Saved
        End If
                    
        ActiveWorkbook.Worksheets("Remove_Duplicates").Range("A1").Resize(UBound(Saved_Column, 1), UBound(Saved_Column, 2)).Value = Saved_Column
                
        ' Remove the duplicates in the current column in order to search for the missing items from the filter
        If ActiveWorkbook.Worksheets("Remove_Duplicates").UsedRange.Rows.Count <> 1 Then
            ActiveWorkbook.Worksheets("Remove_Duplicates").Range(ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(1, 1), ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(ActiveWorkbook.Worksheets("Remove_Duplicates").UsedRange.Rows.Count, 1)).RemoveDuplicates Columns:=1, Header:=xlNo
        End If
        
        Counter_Rows_Saved = 1
        Do While ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(Counter_Rows_Saved, 1) <> ""
            cb_tablename.AddItem (ActiveWorkbook.Worksheets("Remove_Duplicates").Cells(Counter_Rows_Saved, 1))
            
            Counter_Rows_Saved = Counter_Rows_Saved + 1
        Loop
        
        ActiveWorkbook.Worksheets("Remove_Duplicates").Columns(1).EntireColumn.Delete
        
        ActiveWorkbook.Save
        
        If ActiveCell_List_Object Is Nothing Then
            cb_tablename.ListIndex = 0
        Else
            Set Table_Intersect = Intersect(ActiveSheet.ListObjects(ActiveCell_List_Object.Name).Range, Selection)
            
            For i = 0 To cb_tablename.ListCount - 1
                If cb_tablename.List(i) = Table_Intersect.ListObject.DisplayName Then
                    cb_tablename.ListIndex = i
                    
                    Exit For
                End If
            Next i
        End If
        
    End If
    
    View_Filters_Form.Show vbModeless

End Sub


Private Sub cb_tablename_Change()

    lb_filters_table.RowSource = ""
    lb_filters_table.Clear

    Dim ws_saved_fil As Worksheet
    Set ws_saved_fil = ActiveWorkbook.Worksheets("Saved_Filters")

    If cb_tablename.Text = "All" Then
        lb_filters_table.RowSource = "'Saved_Filters'!A2:G" & CStr(ws_saved_fil.UsedRange.Rows.Count)
    ElseIf cb_tablename.Text = "" Then
        Exit Sub
    Else
        Dim Counter_Rows_Saved As Integer
        Counter_Rows_Saved = 2
    
        Dim Counter_Listbox_Items As Integer
        Counter_Listbox_Items = 0
    
        Do While ws_saved_fil.Cells(Counter_Rows_Saved, 1) <> ""
            If ws_saved_fil.Cells(Counter_Rows_Saved, 1) = cb_tablename.Text Then
                lb_filters_table.AddItem ws_saved_fil.Cells(Counter_Rows_Saved, 1)
            
                For i = 1 To 6
                    lb_filters_table.List(Counter_Listbox_Items, i) = ws_saved_fil.Cells(Counter_Rows_Saved, i + 1)
                Next i
            
                Counter_Listbox_Items = Counter_Listbox_Items + 1
            End If
            
            Counter_Rows_Saved = Counter_Rows_Saved + 1
        Loop
        
        If lb_filters_table.ListCount = 1 Then
            lb_filters_table.ListIndex = 0
        Else
            Dim Default_Filter_Index As Integer
            
            Default_Filter_Index = GetSetting("Filter_Save", "Default_Filters", cb_tablename.Text, -1)
            
            If Default_Filter_Index <> -1 Then
                lb_filters_table.ListIndex = Default_Filter_Index
            End If
        End If
    
    End If


End Sub


Private Sub cb_edit_filter_Click()

    If lb_filters_table.ListIndex = -1 Then
        MsgBox ("Please select a filter in order to edit."), vbCritical
        
        Exit Sub
    Else
        For i = 0 To lb_filters_table.ListCount - 1
            If lb_filters_table.Selected(i) = True Then
                Call Save_Filter_Form.Save_Filter_Form_Show(lb_filters_table.List(i, 0), lb_filters_table.List(i, 2), lb_filters_table.List(i, 4), lb_filters_table.List(i, 6))
            End If
        Next i
    End If
    
    View_Filters_Form.cb_tablename.Clear
    View_Filters_Form.lb_filters_table.Clear
    View_Filters_Form.Hide
    
    View_Filters_Show_Form
    

End Sub


Private Sub cb_delete_filter_Click()

    If lb_filters_table.ListIndex = -1 Then
        MsgBox ("Please select a filter in order to delete it."), vbCritical
        
        Exit Sub
    Else
        For i = 0 To lb_filters_table.ListCount - 1
            If lb_filters_table.Selected(i) = True Then
                Dim isDefault As Integer
                
                isDefault = GetSetting("Filter_Save", "Default_Filters", lb_filters_table.List(i, 0), -1)
                
                If isDefault = lb_filters_table.ListIndex Then
                    DeleteSetting "Filter_Save", "Default_Filters", lb_filters_table.List(i, 0)
                End If
                
                Dim Counter_Rows_Saved As Integer
                Counter_Rows_Saved = 2
                
                Do While ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved, 1) <> ""
                    If ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved, 1) = lb_filters_table.List(i, 0) And ActiveWorkbook.Worksheets("Saved_Filters").Cells(Counter_Rows_Saved, 3) = lb_filters_table.List(i, 2) Then
                        ActiveWorkbook.Worksheets("Saved_Filters").Rows(Counter_Rows_Saved).EntireRow.Delete
                        
                        Exit Do
                    End If
                    
                    Counter_Rows_Saved = Counter_Rows_Saved + 1
                Loop
            End If
        Next i
    End If
    
    View_Filters_Form.cb_tablename.Clear
    View_Filters_Form.lb_filters_table.Clear
    View_Filters_Form.Hide
    
    View_Filters_Show_Form
End Sub


Private Sub cb_make_default_Click()

    If cb_tablename.Text = "All" Then
        MsgBox ("Can't make the filter default on all tables"), vbCritical
    Else
        If lb_filters_table.ListIndex = -1 Then
            MsgBox ("Please select a filter in order to make it default."), vbCritical
            
            Exit Sub
        Else
            SaveSetting "Filter_Save", "Default_Filters", cb_tablename.Text, lb_filters_table.ListIndex
        End If
    End If

End Sub

Private Sub cb_apply_filter_Click()
    
    fil_save_Apply_Filter_To_Table
    
End Sub

Private Sub cb_ok_Click()
    
    fil_save_Apply_Filter_To_Table
    
    Unload View_Filters_Form

End Sub

Public Sub fil_save_Apply_Filter_To_Table()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ws_saved_fil As Worksheet
    Set ws_saved_fil = ActiveWorkbook.Worksheets("Saved_Filters")
    
    Dim Counter_Rows_Saved As Integer
    Counter_Rows_Saved = 2
    
    Dim Filters_Selected_Index As Integer
    Filters_Selected_Index = lb_filters_table.ListIndex
    
    If Filters_Selected_Index = -1 Then
        MsgBox ("Please select a filter from the list in order to apply."), vbCritical
        
        Exit Sub
    End If
    
    Dim ActiveCell_List_Object As Variant
    Set ActiveCell_List_Object = ActiveCell.ListObject
    
    If ActiveCell_List_Object Is Nothing Then
        MsgBox ("Please put the cursor in a table to apply the filter to it."), vbCritical
        
        Exit Sub
    End If
    
    Dim Found_Saved_Criteria As String
    
    Do While ws_saved_fil.Cells(Counter_Rows_Saved, 1) <> ""
        If ws_saved_fil.Cells(Counter_Rows_Saved, 1) = lb_filters_table.List(Filters_Selected_Index, 0) And ws_saved_fil.Cells(Counter_Rows_Saved, 3) = lb_filters_table.List(Filters_Selected_Index, 2) Then
            Found_Saved_Criteria = ws_saved_fil.Cells(Counter_Rows_Saved, 8)
            
            Exit Do
        End If
        
        Counter_Rows_Saved = Counter_Rows_Saved + 1
    Loop
    
    Dim Split_Filters As Variant
    
    If InStr(Found_Saved_Criteria, ";") Then
        Split_Filters = Split(Found_Saved_Criteria, ";")
    Else
        Split_Filters = Array(Found_Saved_Criteria)
    End If
    
    
    Dim Header_Intersect_Range As Range
    Set Header_Intersect_Range = ws.ListObjects(ActiveCell.ListObject.Name).HeaderRowRange
    
    If ws.FilterMode = True Then
        ActiveCell_List_Object.AutoFilter.ShowAllData
        Header_Intersect_Range.AutoFilter
    Else
        Header_Intersect_Range.AutoFilter
    End If
    
    Dim Table_Intersect As Variant
    Set Table_Intersect = Intersect(ws.ListObjects(ActiveCell_List_Object.Name).Range, Selection)
    
    Dim Filter As Variant
    Dim Split_Header_Filter As Variant
    
    Dim Sign As String
    Dim Found_Column As Integer
    Dim Split_Multiple_Criteria As Variant
    
    Dim Dict_Criterias As Scripting.Dictionary
    Dim Array_Criterias As Variant
    Dim Saved_Column As Variant
    
    Dim Duplicates_Ws_Last_Row As Integer
    
    For Each Filter In Split_Filters
        Split_Header_Filter = Split(Filter, ":")
        
        Found_Column = Header_Intersect_Range.Find(Split_Header_Filter(0)).Column
        
        If InStr(Split_Header_Filter(1), "|") Then
            Split_Multiple_Criteria = Split(Split_Header_Filter(1), "|")
        Else
            Split_Multiple_Criteria = CStr(Split_Header_Filter(1))
        End If
        
        Header_Intersect_Range.AutoFilter field:=Found_Column - Header_Intersect_Range.Column + 1, Operator:=xlFilterValues, Criteria1:=Split_Multiple_Criteria
    Next Filter

End Sub

Private Sub UserForm_Click()

End Sub
