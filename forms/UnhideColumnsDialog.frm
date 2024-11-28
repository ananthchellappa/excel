VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnhideColumnsDialog 
   Caption         =   "Choose Cols to Unhide"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3630
   OleObjectBlob   =   "UnhideColumnsDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnhideColumnsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private isOKPressed As Boolean ' Flag to indicate whether OK was pressed

Private Sub UserForm_Initialize()
    Dim ActiveTable As ListObject
    Dim ColumnHeader As Range
    Dim SelectedRange As Range
    Dim ColumnIndex As Integer
    Dim ColumnLetter As String
    Dim StraddledHiddenIndices As Collection
    Dim i As Integer

    On Error GoTo NotInTable
    ' Check if active cell is in a table
    Set ActiveTable = ActiveCell.ListObject
    On Error GoTo 0

    ' Initialize variables
    Set SelectedRange = Selection
    Set StraddledHiddenIndices = New Collection

    ' Populate the ListBox with all hidden column names with Excel column letters as prefix
    For Each ColumnHeader In ActiveTable.HeaderRowRange
        If ColumnHeader.EntireColumn.Hidden Then
            ' Get the Excel column letter
            ColumnLetter = Split(ColumnHeader.Address, "$")(1)
            
            ' Add the prefixed column name to the ListBox
            Me.lstHiddenColumns.AddItem ColumnLetter & " - " & ColumnHeader.Value

            ' Check if the hidden column is straddled by the user's selection
            ColumnIndex = ColumnHeader.Column
            If Not Intersect(ActiveTable.Range.Columns(ColumnIndex), SelectedRange) Is Nothing Then
                StraddledHiddenIndices.Add Me.lstHiddenColumns.ListCount - 1 ' Store the index
            End If
        End If
    Next ColumnHeader

    ' Scroll the ListBox to the first straddled hidden column, if any
    If StraddledHiddenIndices.count > 0 Then
        Me.lstHiddenColumns.TopIndex = StraddledHiddenIndices(1)
    End If
    Exit Sub

NotInTable:
    MsgBox "Please place the cursor in an Excel table and run the subroutine.", vbCritical, "Error"
    Unload Me
End Sub

Private Sub chkSelectAll_Click()
    Dim i As Integer

    ' Toggle selection of all ListBox items based on the checkbox state
    If Me.chkSelectAll.Value = True Then
        For i = 0 To Me.lstHiddenColumns.ListCount - 1
            Me.lstHiddenColumns.Selected(i) = True
        Next i
    Else
        For i = 0 To Me.lstHiddenColumns.ListCount - 1
            Me.lstHiddenColumns.Selected(i) = False
        Next i
    End If
End Sub

Private Sub btnOK_Click()
    ' Set the flag to indicate OK was pressed
    isOKPressed = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    ' Hide the form instead of unloading it
    Me.Hide
End Sub

' Public function to check if OK was pressed
Public Function OKPressed() As Boolean
    OKPressed = isOKPressed
End Function

