Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim stopAnimation As Boolean
Dim lastColors As Object ' Dictionary to store the last color of each blank cell

Sub AnimateBlankCells()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim startTime As Double
    Dim duration As Double
    Dim rStart As Integer, rEnd As Integer
    Dim gStart As Integer, gEnd As Integer
    Dim bStart As Integer, bEnd As Integer
    Dim stepValue As Integer
    Dim randomDelay As Integer
    Dim blankCells As Collection
    Dim i As Integer
    Dim elapsedTime As Double
    Dim newColor As Long
    Dim style As Integer
    Dim rowOffset As Integer, colOffset As Integer
    Dim LL As Range, UR As Range
    Dim lowerLeft As String, upperRight As String
    Dim firstRow As Integer, lastRow As Integer, firstCol As Integer, lastCol As Integer
    
    ' Set worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize dictionary to store last colors
    Set lastColors = CreateObject("Scripting.Dictionary")
    
    ' Retrieve range from B10 and B11 if provided
    lowerLeft = ws.Range("B10").value
    upperRight = ws.Range("B11").value
    
    If lowerLeft <> "" And upperRight <> "" Then
        On Error Resume Next
        Set LL = ws.Range(lowerLeft)
        Set UR = ws.Range(upperRight)
        On Error GoTo 0
        
        If Not LL Is Nothing And Not UR Is Nothing Then
            ' Define the range using the provided coordinates
            firstRow = UR.Row
            lastRow = LL.Row
            firstCol = LL.Column
            lastCol = UR.Column
            Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
        Else
            MsgBox "Invalid range coordinates in B10 or B11. Using selected range instead.", vbExclamation
            Set rng = Selection
        End If
    Else
        ' Use selected range if B10 or B11 are empty
        Set rng = Selection
    End If

    ' Validate color range values
    rStart = GetValidatedNumber(ws.Range("B1").value, 0, 255)
    rEnd = GetValidatedNumber(ws.Range("B2").value, 0, 255)
    gStart = GetValidatedNumber(ws.Range("B3").value, 0, 255)
    gEnd = GetValidatedNumber(ws.Range("B4").value, 0, 255)
    bStart = GetValidatedNumber(ws.Range("B5").value, 0, 255)
    bEnd = GetValidatedNumber(ws.Range("B6").value, 0, 255)
    duration = GetValidatedNumber(ws.Range("B7").value, 1, 100)

    stepValue = GetValidatedNumber(ws.Range("B8").value, 1, 255)
    style = GetValidatedNumber(ws.Range("B9").value, 0, 8)

    ' Define movement offsets
    Dim movementOffsets As Variant
    movementOffsets = Array( _
        Array(0, 0), Array(-1, 0), Array(0, 1), Array(1, 0), _
        Array(0, -1), Array(-1, -1), Array(-1, 1), Array(1, 1), Array(1, -1))
    
    rowOffset = movementOffsets(style)(0)
    colOffset = movementOffsets(style)(1)

    ' Collect blank cells
    Set blankCells = New Collection
    For Each cell In rng
        If cell.value = "" Then
            blankCells.Add cell
            lastColors(cell.Address) = RGB( _
                GetSteppedValue(rStart, rEnd, stepValue), _
                GetSteppedValue(gStart, gEnd, stepValue), _
                GetSteppedValue(bStart, bEnd, stepValue))
        End If
    Next cell

    ' If no blank cells, exit
    If blankCells.Count = 0 Then
        MsgBox "No blank cells found in the selected range.", vbExclamation
        Exit Sub
    End If

    ' Start animation
    stopAnimation = False
    startTime = Timer

    Do
        elapsedTime = Timer - startTime
        If elapsedTime >= duration Or stopAnimation Then Exit Do
        
        Dim tempColors As Object
        Set tempColors = CreateObject("Scripting.Dictionary")

        If style = 0 Then ' No sliding, update every cycle
            For Each cell In blankCells
                tempColors(cell.Address) = RGB( _
                    GetSteppedValue(rStart, rEnd, stepValue), _
                    GetSteppedValue(gStart, gEnd, stepValue), _
                    GetSteppedValue(bStart, bEnd, stepValue))
            Next cell
        Else ' Apply sliding effect
            For Each cell In blankCells
                Dim prevRow As Integer, prevCol As Integer
                Dim prevCell As Range
                
                prevRow = cell.Row + rowOffset
                prevCol = cell.Column + colOffset
                
                On Error Resume Next
                Set prevCell = ws.Cells(prevRow, prevCol)
                On Error GoTo 0
                
                If Not prevCell Is Nothing And lastColors.Exists(prevCell.Address) Then
                    tempColors(cell.Address) = lastColors(prevCell.Address)
                Else
                    tempColors(cell.Address) = RGB( _
                        GetSteppedValue(rStart, rEnd, stepValue), _
                        GetSteppedValue(gStart, gEnd, stepValue), _
                        GetSteppedValue(bStart, bEnd, stepValue))
                End If
            Next cell
        End If

        ' Apply colors
        For Each cell In blankCells
            If tempColors.Exists(cell.Address) Then
                cell.Interior.Color = tempColors(cell.Address)
                lastColors(cell.Address) = tempColors(cell.Address)
            End If
        Next cell

        ' ? Force Excel to refresh display after each frame
        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False

        randomDelay = 80 'RandBetween(100, 200)
        Sleep randomDelay
        
    Loop

    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    'MsgBox "Animation complete! Colors stored.", vbInformation
End Sub

' Helper Functions
Function GetValidatedNumber(value As Variant, min As Integer, max As Integer) As Integer
    If IsNumeric(value) Then
        GetValidatedNumber = Application.WorksheetFunction.max(min, Application.WorksheetFunction.min(max, CInt(value)))
    Else
        GetValidatedNumber = min ' Default to min value if input is invalid
    End If
End Function

Function GetSteppedValue(lower As Integer, upper As Integer, stepSize As Integer) As Integer
    Dim value As Integer
    If stepSize = 1 Then
        GetSteppedValue = RandBetween(lower, upper)
        Exit Function
    End If
    value = RandBetween(lower, upper)
    value = Int(value / stepSize) * stepSize
    If value < lower Then value = lower
    If value > upper Then value = upper
    GetSteppedValue = value
End Function

Function RandBetween(lower As Integer, upper As Integer) As Integer
    RandBetween = Int((upper - lower + 1) * Rnd + lower)
End Function


' Stop function
Sub StopBlankCellAnimation()
    stopAnimation = True
End Sub



