VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HideColumnsDialog 
   Caption         =   "Choose Columns to Hide"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
   OleObjectBlob   =   "HideColumnsDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HideColumnsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isOKPressed As Boolean ' Flag to indicate whether OK was pressed

Private Sub UserForm_Initialize()
    Dim ActiveTable As ListObject
    Dim ColumnHeader As Range
    Dim ColumnLetter As String

    On Error GoTo NotInTable
    ' Check if active cell is in a table
    Set ActiveTable = ActiveCell.ListObject
    On Error GoTo 0

    ' Initialize the ListBox with visible column names prefixed by Excel column letters
    For Each ColumnHeader In ActiveTable.HeaderRowRange
        If Not ColumnHeader.EntireColumn.Hidden Then
            ' Get the Excel column letter
            ColumnLetter = Split(ColumnHeader.Address, "$")(1)
            
            ' Add the prefixed column name to the ListBox
            Me.lstColumnNames.AddItem ColumnLetter & " - " & ColumnHeader.Value
        End If
    Next ColumnHeader

    ' If no visible columns, close the form silently
    If Me.lstColumnNames.ListCount = 0 Then
        Unload Me
        Exit Sub
    End If
    Exit Sub

NotInTable:
    MsgBox "Please place the cursor in an Excel table and run the subroutine.", vbCritical, "Error"
    Unload Me
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

