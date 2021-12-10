Attribute VB_Name = "Module6"
Option Explicit

Sub DeleteThis()
    Application.DisplayAlerts = False
    'ThisWorkbook.Sheets(ActiveSheet.Index).Delete  ' this will try to delete the sheet in the Personal Macro Workbook
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub

Sub gen_Legend()
Attribute gen_Legend.VB_Description = "create (if not existing) a new Legend sheet and display the data from selected row in currently viewed table"
Attribute gen_Legend.VB_ProcData.VB_Invoke_Func = "X\n14"


Application.ScreenUpdating = False
 

Dim isheet, column_name As String
Dim irow, irow_header, lastcolumn, lastrow, counter As Double
Dim table_name, index_column, header_column, value_column, blank_column As String
Dim i, j As Integer


'name of sheet with source data
isheet = ActiveSheet.Name
'row number from where header has to be copied
irow_header = Range(ActiveCell.ListObject.HeaderRowRange.Address).Row
'row number from where data has to be copied
irow = ActiveCell.Row
'total number of columns in source data
lastcolumn = Sheets(isheet).Range("zz" & irow_header).End(xlToLeft).Column
'calculates number of tables required, one table is required for every 20 columns
counter = Application.WorksheetFunction.RoundUp(lastcolumn / 20, 0)

'check if Legend sheet is present
If Evaluate("ISREF('" & "Legend" & "'!A1)") = False Then
 'if its not present then adds new sheet with that name and row 3 is assigned as lastrow (where data will be added)
 Sheets.Add
 ActiveSheet.Name = "Legend"
 lastrow = 3
Else
'if Legend is already present, then 5 rows after last row of table is assigned as lastrow
lastrow = Sheets("Legend").Range("B10000").End(xlUp).Row + 5
End If

 Sheets("Legend").Select
 Sheets("Legend").Columns("A").ColumnWidth = 8.43

' this is used to add data from source to legend sheet. it uses loop to add data in each table. 20 columns per table
For i = 1 To counter

'column names for different values of table
index_column = ColumnLetter(2 + (i - 1) * 4)
header_column = ColumnLetter(3 + (i - 1) * 4)
value_column = ColumnLetter(4 + (i - 1) * 4)
blank_column = ColumnLetter(5 + (i - 1) * 4)

'table creation
    table_name = "t_" & i & Format(Now(), "yyyymmdd_hhmmss")
    Sheets("Legend").ListObjects.Add(xlSrcRange, Range(index_column & lastrow - 1 & ":" & value_column & lastrow + 18), , xlNo).Name = table_name
    Sheets("Legend").Range(table_name & "[#All]").Select
    Sheets("Legend").ListObjects(table_name).TableStyle = "TableStyleMedium15"
    Sheets("Legend").ListObjects(table_name).ShowHeaders = False

'adding row numbers
For j = 1 To 20
Sheets("Legend").Cells(lastrow + j - 1, 2 + (i - 1) * 4).Value = (i - 1) * 20 + j
Next j




' this part copy pastes column headers
Sheets(isheet).Range(ColumnLetter(3 + (i - 1) * 20) & irow_header & ":" & ColumnLetter(2 + 20 + (i - 1) * 20) & irow_header).Copy
Sheets("Legend").Range(header_column & lastrow).PasteSpecial Paste:=xlValue, Transpose:=True

' this part copy pastes data of specific row
Sheets(isheet).Range(ColumnLetter(3 + (i - 1) * 20) & irow & ":" & ColumnLetter(2 + 20 + (i - 1) * 20) & irow).Copy
Sheets("Legend").Range(value_column & lastrow).PasteSpecial Paste:=xlValue, Transpose:=True

 'column autofit
 Sheets("Legend").Columns(index_column).EntireColumn.AutoFit
 Sheets("Legend").Columns(header_column).EntireColumn.AutoFit
 Sheets("Legend").Columns(value_column).EntireColumn.AutoFit

'column width if auto fit has reduced below minimum required value
 If Sheets("Legend").Columns(index_column).ColumnWidth < 5 Then Sheets("Legend").Columns(index_column).ColumnWidth = 5
 If Sheets("Legend").Columns(header_column).ColumnWidth < 10 Then Sheets("Legend").Columns(header_column).ColumnWidth = 10
 If Sheets("Legend").Columns(value_column).ColumnWidth < 10 Then Sheets("Legend").Columns(value_column).ColumnWidth = 10
 Sheets("Legend").Columns(blank_column).ColumnWidth = 4
 
 'column allingment
 Sheets("Legend").Columns(index_column & ":" & value_column).HorizontalAlignment = xlCenter

Next i


' adding macro button assigning macro to it
 Sheets("Legend").Buttons.Add(Range("G" & lastrow + 21).Left, Range("G" & lastrow + 21).Top, 93.75, 27).Select
    Selection.OnAction = "DeleteThis"
    Selection.Characters.Text = "CLOSE"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With


'this brings focus to new table
Application.GoTo Worksheets("Legend").Range("A" & lastrow), True

Application.ScreenUpdating = True

End Sub

'function that converts column number to column name
Function ColumnLetter(ByVal ColumnN As Integer) As String
  ColumnLetter = Split(Cells(1, ColumnN).Address, "$")(1)
End Function


Sub incr_font()
    Selection.Font.Size = Selection.Font.Size + 1
End Sub

Sub decr_font()
    Selection.Font.Size = Selection.Font.Size - 1
End Sub
