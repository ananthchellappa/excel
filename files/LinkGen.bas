Attribute VB_Name = "LinkGen"
Option Explicit
Dim arrData
Private Sub forSingleCell()
    createLink ActiveCell
End Sub


Public Sub forSelection()
    Dim ws As Worksheet
    Dim r As Range
    
    Application.ScreenUpdating = False
    For Each r In Selection.Cells
        If isReference(r) Then
            createLink r
        End If
    Next r
    Application.ScreenUpdating = True
End Sub

Private Sub createLink(r As Range)
    Dim rTemp As Range
    Set rTemp = ActiveSheet.Cells(r.Row, ActiveSheet.UsedRange.Columns.Count + ActiveSheet.UsedRange.Cells(1, 1).Column + 1).Resize(r.MergeArea.Rows.Count, r.MergeArea.Columns.Count)
    r.MergeArea.Copy
    rTemp.PasteSpecial xlPasteFormats
    r.Parent.Hyperlinks.Add Anchor:=r, Address:="", SubAddress:=Replace(r.Formula, "=", "")
    rTemp.Copy
    r.PasteSpecial xlPasteFormats
    rTemp.EntireColumn.Delete
    Application.CutCopyMode = False
    
End Sub

Private Function isReference(r)
    Dim f
    Dim arr
    Dim destRange As Range
    If Left(r.Formula, 1) <> "=" Then Exit Function
    f = Replace(r.Formula, "=", "")
    arr = Split(f, "!")
    On Error Resume Next
    Set destRange = Sheets(Replace(arr(0), "'", "")).Range(arr(1))
    On Error GoTo 0
    If Not destRange Is Nothing Then isReference = True
End Function
Private Function isReferenceStr(f As Variant)
    Dim arr
    Dim destRange As Range
    If Left(f, 1) <> "=" Then Exit Function
    f = Replace(f, "=", "")
    arr = Split(f, "!")
    On Error Resume Next
    Set destRange = Sheets(Replace(arr(0), "'", "")).Range(arr(1))
    On Error GoTo 0
    If Not destRange Is Nothing Then isReferenceStr = True
End Function


Public Sub forEntireWorksheet()
    Dim ws As Worksheet
    Dim r As Range
    Dim rTemp As Range
    Dim i, j
    Application.ScreenUpdating = False
    arrData = ActiveSheet.UsedRange.Formula
    For i = LBound(arrData) To UBound(arrData)
        For j = LBound(arrData, 2) To UBound(arrData, 2)
            If isReferenceStr(arrData(i, j)) Then
                Set r = ActiveSheet.UsedRange.Cells(i, j)
                createLink r
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub

