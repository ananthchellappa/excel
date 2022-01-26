Public Type commandItem
    cell As Range
    cellValue As String
    command As String
    commandText As String
    commandDrawn As String
    customWidth As String
End Type

Sub createWave()

Dim globalLineWidth As Integer
Dim numberIterations, globalTransitionsCellWidth As Integer
Dim allTransitionsMode As Boolean
Dim logicCommands() As commandItem
Dim c As Range
Dim addNewRows As Boolean

'Set mode
addNewRows = True

'Checks
If Selection.Rows.Count > 1 Then
    MsgBox "Selecting more than 1 row is not allowed.", vbInformation
    Exit Sub
End If

'Check for at least one valid logic command
IsValid = False
For Each c In Selection
    If Left(Trim(UCase(c.Value)), 1) = "W" Or _
       Left(Trim(UCase(c.Value)), 1) = "M" Or _
       Left(Trim(UCase(c.Value)), 1) = "E" Or _
       Left(Trim(UCase(c.Value)), 2) = "AT" Or _
       Left(Trim(UCase(c.Value)), 2) = "AX" Then
    Else
        IsValid = True  'if there is at least one logic command found
    End If
Next c

If IsValid = False Then
    MsgBox "Selection not valid, no logic commands.", vbInformation
    Exit Sub
End If

For Each c In Selection
    If Trim(c) <> "" Then
        commandsFirstColWithValue = c.Column
        Exit For
    End If
Next c

commandsLastCol = ActiveSheet.Cells(Selection.Row, 16384).End(xlToLeft).Column
If Selection.Address <> Selection.EntireRow.Address Then
    commandsLastCol = Selection.Columns(Selection.Columns.Count).Column
Else
    numberIterations = 10 'default for M if not provided and entire row selected
End If

If IsEmpty(commandsFirstColWithValue) Or IsEmpty(commandsLastCol) Then
    MsgBox "No commands selected.", vbInformation
    Exit Sub
End If

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Set commandsRange = ActiveSheet.Range(Cells(Selection.Row, commandsFirstColWithValue), Cells(Selection.Row, commandsLastCol))

If addNewRows = True Then
    Set newRowsRange = ActiveSheet.Range(Cells(commandsRange.Row + 1, 1), Cells(commandsRange.Row + 3, 1)).EntireRow
    newRowsRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newRowsRange = ActiveSheet.Range(Cells(commandsRange.Row + 1, 1), Cells(commandsRange.Row + 3, 1)).EntireRow
    newRowsRange.ClearFormats
End If

'Debug.Print commandsRange.Address

'Get global commands from selection
' W1, W2, W3 -> Width of the line
' Mxx -> Number of Cells, default 10
' ATxx (xx optional) -> All transitions AND width (default same as logic cells
' AXxx (xx mandatory) -> All bus transitions AND width
' E (explicit) -> Retain previous logic value

globalLineWidth = 2 'default for W if not provided
globalTransitionsCellWidth = 100 'default if no number was provided with command AT
allTransitionsMode = False
explicitMode = False

missingM = True
If Selection.Address <> Selection.EntireRow.Address Then 'Not entire row is selected
    missingM = False
End If

For Each c In commandsRange

    isGlobalCommand = False

    If Left(Trim(UCase(c.Value)), 1) = "W" Then
        checkValue = Replace(Trim(c.Value), "W", "")
        Select Case checkValue
            Case 1, 2, 3
                globalLineWidth = CInt(checkValue)
                isGlobalCommand = True
            Case Else
                MsgBox (c.Value & " is not a valid W command for line thickness."), vbInformation
                Exit Sub
        End Select
    End If
        
    If Left(Trim(UCase(c.Value)), 1) = "M" Then
        checkValue = Replace(Trim(c.Value), "M", "")
        Select Case checkValue
            Case ""
                MsgBox "M command needs to be provided with a number but is empty.", vbInformation
                Exit Sub
            Case Else
                If IsNumeric(checkValue) Then
                    numberIterations = CInt(checkValue)
                    isGlobalCommand = True
                    missingM = False
                Else
                    MsgBox (c.Value & " is not a valid M command for number of iterations."), vbInformation
                    Exit Sub
                End If
        End Select
    End If
    
    If Left(Trim(UCase(c.Value)), 2) = "AT" Or Left(Trim(UCase(c.Value)), 2) = "AX" Then
    
        allTransitionsMode = True
        isGlobalCommand = True
    
        checkValue = Replace(Trim(c.Value), "AT", "")
        checkValue = Replace(checkValue, "AX", "")

        If IsNumeric(checkValue) Then
            globalTransitionsCellWidth = CInt(checkValue)
        End If

    End If
    
    If Left(Trim(UCase(c.Value)), 1) = "E" Then
        explicitMode = True
        isGlobalCommand = True
    End If
    
    If isGlobalCommand = True Then
        lastColWithGlobalCommand = c.Column
    End If
    
Next c

If IsEmpty(lastColWithGlobalCommand) Then
    lastColWithGlobalCommand = commandsRange.Column - 1
End If

commandsRange.EntireRow.Font.ColorIndex = xlAutomatic

numberIterations = Application.WorksheetFunction.Max(numberIterations, commandsLastCol - (lastColWithGlobalCommand))

'Get logic commands
Set logicCommandsRange = ActiveSheet.Range(Cells(commandsRange.Row, lastColWithGlobalCommand + 1), Cells(commandsRange.Row, lastColWithGlobalCommand + numberIterations))

i = 0
For Each c In logicCommandsRange

    ReDim Preserve logicCommands(i)
    Set logicCommands(i).cell = c
      
    logicCommands(i).cellValue = c.Value
    
    If Left(UCase(c.Value), 1) = "=" Then
        logicCommands(i).command = "="
        logicCommands(i).commandText = Trim(Replace(c.Value, "=", ""))
    End If
    
    If Left(CStr(c.Value), 1) = "1" Then
        logicCommands(i).command = "1"
    End If
    
    If Left(CStr(c.Value), 1) = "0" Then
        logicCommands(i).command = "0"
    End If
    
    If Left(UCase(c.Value), 1) = "X" Then
        logicCommands(i).command = "X"
        
        checkWidth = Replace(UCase(c.Value), "X", "")
        If IsNumeric(checkWidth) Then
            logicCommands(i).customWidth = checkWidth
        End If
    End If
    
    If Left(UCase(c.Value), 1) = "T" Then
        logicCommands(i).command = "T"
        
        checkWidth = Replace(UCase(c.Value), "T", "")
        If IsNumeric(checkWidth) Then
            logicCommands(i).customWidth = checkWidth
        End If
    End If
    
    If Left(UCase(c.Value), 1) = "U" Then
        logicCommands(i).command = "U"
    End If
    
    i = i + 1

Next c

'Check if bus and if yes, bus wins
isBus = False
For i = 0 To UBound(logicCommands)
    If InStr(logicCommands(i).command, "=") > 0 Then
        isBus = True
        Exit For
    End If
Next i

'Remove and mark invalid commands
For i = 0 To UBound(logicCommands)

    If isBus = True Then 'Remove non bus transision commands
        Select Case True
            Case logicCommands(i).command <> "X" And InStr(logicCommands(i).command, "=") = 0 And logicCommands(i).command <> "U" And Trim(logicCommands(i).cell.Value2) <> ""
                logicCommands(i).cell.Font.Color = -16776961
                logicCommands(i).command = ""
                logicCommands(i).commandText = ""
                logicCommands(i).customWidth = ""
        End Select
    Else
        Select Case True
            Case logicCommands(i).command <> "1" And logicCommands(i).command <> "0" And logicCommands(i).command <> "T" And logicCommands(i).command <> "U" And Trim(logicCommands(i).cell.Value2) <> ""
                logicCommands(i).cell.Font.Color = -16776961
                logicCommands(i).command = ""
                logicCommands(i).commandText = ""
                logicCommands(i).customWidth = ""
        End Select
    End If

Next i

' U -> Unkown pattern
' '= / =VALUE -> bus command and value as text in cell
' X -> bus transition
' Txx (xx optional= -> transition

'Fill first column
logicCommands(0).cell.Offset(2, 0).EntireRow.ClearFormats
logicCommands(0).cell.Offset(2, 0).EntireRow.ClearContents

Select Case True
    Case logicCommands(0).command = "" Or logicCommands(0).command = "0"
        If isBus = False Then
            Call drawBottom(logicCommands(0), globalLineWidth, False)
            If explicitMode = True Then
                nextCommand = "Bottom"
            ElseIf allTransitionsMode = True Then
                nextCommand = "TransitionBottomTop"
            Else
                nextCommand = "Top"
            End If
        Else
            Call drawBus(logicCommands(0), globalLineWidth, False, logicCommands(0).commandText)
            nextCommand = "Bus"
        End If
        
    Case logicCommands(0).command = "1"
        If isBus = False Then
            Call drawTop(logicCommands(0), globalLineWidth, False)
            If explicitMode = True Then
                nextCommand = "Top"
            ElseIf allTransitionsMode = True Then
                nextCommand = "TransitionTopBottom"
            Else
                nextCommand = "Bottom"
            End If
        Else
            Call drawBus(logicCommands(0), globalLineWidth, False, logicCommands(0).commandText)
            nextCommand = "Bus"
        End If
        
    Case logicCommands(0).command = "T"
        Call drawTransitionBottomTop(logicCommands(0), globalLineWidth, False)
        nextCommand = "Top"
    Case logicCommands(0).command = "="
        Call drawBus(logicCommands(0), globalLineWidth, False, logicCommands(0).commandText)
        nextCommand = "Bus"
    Case logicCommands(0).command = "X"
        Call drawBusTransition(logicCommands(0), globalLineWidth, False)
        nextCommand = "Bus"
    Case logicCommands(0).command = "U"
        Call drawU(logicCommands(0), globalLineWidth, False)
        nextCommand = "U"
End Select


'Fill iteration columns
For i = 1 To UBound(logicCommands) - 1

'Get latest logic command drawn
For u = i - 1 To 0 Step -1
    If logicCommands(u).commandDrawn <> "Unknown" Then
        lastCommandDrawn = logicCommands(u).commandDrawn
        Exit For
    Else
        lastCommandDrawn = "Bottom" 'Defaults to 0 same as for wave start
    End If
Next u

Debug.Print lastCommandDrawn

    Select Case True
        
        Case logicCommands(i).command = ""
            If isBus = False Then
                If nextCommand = "TransitionBottomTop" Then
                    Call drawTransitionBottomTop(logicCommands(i), globalLineWidth, False)
                    nextCommand = "Top"
                    GoTo ExitSelect
                End If
                If nextCommand = "TransitionTopBottom" Then
                    Call drawTransitionTopBottom(logicCommands(i), globalLineWidth, False)
                    nextCommand = "Bottom"
                    GoTo ExitSelect
                End If
                If nextCommand = "Top" Then
                    Call drawTop(logicCommands(i), globalLineWidth, False)
                    If explicitMode = True Then
                        nextCommand = "Top"
                    ElseIf allTransitionsMode = True Then
                        nextCommand = "TransitionTopBottom"
                        GoTo ExitSelect
                    Else
                        nextCommand = "Bottom"
                        GoTo ExitSelect
                    End If
                End If
                If nextCommand = "Bottom" Then
                    Call drawBottom(logicCommands(i), globalLineWidth, False)
                    If explicitMode = True Then
                        nextCommand = "Bottom"
                    ElseIf allTransitionsMode = True Then
                        nextCommand = "TransitionBottomTop"
                        GoTo ExitSelect
                    Else
                        nextCommand = "Top"
                        GoTo ExitSelect
                    End If
                End If
            Else
                Call drawBus(logicCommands(i), globalLineWidth, False, logicCommands(i).commandText)
                nextCommand = "Bus"
                GoTo ExitSelect
            End If
        
        Case logicCommands(i).command = "T"
        
            If logicCommands(i - 1).command = "U" Then
                'Draw corresponding connection out from unkown state
                If lastCommandDrawn = "Top" Or lastCommandDrawn = "TransitionBottomTop" Then
                    Call drawTransitionTopBottom(logicCommands(i), globalLineWidth, True)
                    nextCommand = "Bottom"
                End If
                If lastCommandDrawn = "Bottom" Or lastCommandDrawn = "TransitionTopBottom" Then
                    Call drawTransitionBottomTop(logicCommands(i), globalLineWidth, True)
                    nextCommand = "Top"
                End If
        
                GoTo ExitSelect
            End If
        
            If logicCommands(i - 1).command = "0" Then
                Call drawTransitionBottomTop(logicCommands(i), globalLineWidth, False)
                nextCommand = "Top"
                GoTo ExitSelect
            End If
            If logicCommands(i - 1).command = "1" Then
                Call drawTransitionTopBottom(logicCommands(i), globalLineWidth, False)
                nextCommand = "Bottom"
                GoTo ExitSelect
            End If
            
            If nextCommand = "TransitionBottomTop" Then
                Call drawTransitionBottomTop(logicCommands(i), globalLineWidth, False)
                nextCommand = "Top"
                GoTo ExitSelect
            End If
            If nextCommand = "TransitionTopBottom" Then
                Call drawTransitionTopBottom(logicCommands(i), globalLineWidth, False)
                nextCommand = "Bottom"
                GoTo ExitSelect
            End If
            If nextCommand = "Top" Then
                Call drawTransitionTopBottom(logicCommands(i), globalLineWidth, False)
                nextCommand = "Bottom"
                GoTo ExitSelect
            End If
            If nextCommand = "Bottom" Then
                Call drawTransitionBottomTop(logicCommands(i), globalLineWidth, False)
                nextCommand = "Top"
                GoTo ExitSelect
            End If
            
        Case logicCommands(i).command = "1"
            If isBus = False Then
                Call drawTop(logicCommands(i), globalLineWidth, False)
                If explicitMode = True Then
                    nextCommand = "Top"
                ElseIf allTransitionsMode = True Then
                    nextCommand = "TransitionTopBottom"
                Else
                    nextCommand = "Bottom"
                End If
            End If

        Case logicCommands(i).command = "0"
            If isBus = False Then
                Call drawBottom(logicCommands(i), globalLineWidth, False)
                If explicitMode = True Then
                    nextCommand = "Bottom"
                ElseIf allTransitionsMode = True Then
                    nextCommand = "TransitionBottomTop"
                Else
                    nextCommand = "Top"
                End If
            End If

        Case logicCommands(i).command = "="
            Call drawBus(logicCommands(i), globalLineWidth, False, logicCommands(i).commandText)
            nextCommand = "Bus"
        
        Case logicCommands(i).command = "X"
            Call drawBusTransition(logicCommands(i), globalLineWidth, False)
            GoTo ExitSelect
        
        Case logicCommands(i).command = "U"
            Call drawU(logicCommands(i), globalLineWidth, False)
            GoTo ExitSelect
            
ExitSelect:
    End Select

    'Set explicit exceptions
    If allTransitionsMode = True And explicitMode = True Then

        If logicCommands(i).command = "" And logicCommands(i).commandDrawn = "Top" And logicCommands(i + 1).command = "0" Then
            Call drawTransitionTopBottom(logicCommands(i), globalLineWidth, False)
        End If
        
        If logicCommands(i).command = "" And logicCommands(i).commandDrawn = "Bottom" And logicCommands(i + 1).command = "1" Then
            Call drawTransitionBottomTop(logicCommands(i), globalLineWidth, False)
        End If
    End If

Next i

'Fill last column

Select Case True
    Case logicCommands(UBound(logicCommands)).cell = ""
        If isBus = False Then
            If nextCommand = "TransitionBottomTop" Then
                Call drawTransitionBottomTop(logicCommands(UBound(logicCommands)), globalLineWidth, False)
                nextCommand = ""
            End If
            If nextCommand = "TransitionTopBottom" Then
                Call drawTransitionTopBottom(logicCommands(UBound(logicCommands)), globalLineWidth, False)
                nextCommand = ""
            End If
            If nextCommand = "Top" Then
                Call drawTop(logicCommands(UBound(logicCommands)), globalLineWidth, False)
                nextCommand = ""
            End If
            If nextCommand = "Bottom" Then
                Call drawBottom(logicCommands(UBound(logicCommands)), globalLineWidth, False)
                nextCommand = ""
            End If
        Else
            Call drawBus(logicCommands(UBound(logicCommands)), globalLineWidth, False, logicCommands(UBound(logicCommands)).commandText)
            nextCommand = ""
        End If
        
    Case logicCommands(UBound(logicCommands)).command = "T"
        Call drawTransitionBottomTop(logicCommands(UBound(logicCommands)), globalLineWidth, False)
        nextCommand = ""
    Case logicCommands(UBound(logicCommands)).command = "1"
        Call drawTop(logicCommands(UBound(logicCommands)), globalLineWidth, False)
        nextCommand = ""
    Case logicCommands(UBound(logicCommands)).command = "0"
        Call drawBottom(logicCommands(UBound(logicCommands)), globalLineWidth, False)
        nextCommand = ""
    Case logicCommands(UBound(logicCommands)).command = "="
        Call drawBus(logicCommands(i), globalLineWidth, False, logicCommands(UBound(logicCommands)).commandText)
        nextCommand = ""
    Case logicCommands(UBound(logicCommands)).command = "X"
        Call drawBusTransition(logicCommands(UBound(logicCommands)), globalLineWidth, False)
        nextCommand = ""
    Case logicCommands(UBound(logicCommands)).command = "U"
        Call drawU(logicCommands(UBound(logicCommands)), globalLineWidth, False)
        nextCommand = ""
End Select

'Add error in case of missing M command
If missingM = True Then
    logicCommands(UBound(logicCommands)).cell.Offset(2, 1).Value2 = "err: M"
End If

'Mask unknown command chain
For i = 1 To UBound(logicCommands)
    If logicCommands(i - 1).commandDrawn = "Unknown" And logicCommands(i).command = "" Then
        Call drawU(logicCommands(i), globalLineWidth, False, False)
    End If
Next i

'Add connectors
If isBus = False Then
    For i = 1 To UBound(logicCommands)
        If logicCommands(i).commandDrawn = "Top" And logicCommands(i - 1).commandDrawn = "Bottom" Then
            Call drawLeft(logicCommands(i), globalLineWidth, True, True)
        End If
        If logicCommands(i).commandDrawn = "Bottom" And logicCommands(i - 1).commandDrawn = "Top" Then
            Call drawLeft(logicCommands(i), globalLineWidth, True, True)
        End If

        If logicCommands(i).commandDrawn = "Top" And logicCommands(i - 1).commandDrawn = "TransitionTopBottom" Then
            Call drawLeft(logicCommands(i), globalLineWidth, True, True)
        End If
        If logicCommands(i).commandDrawn = "Bottom" And logicCommands(i - 1).commandDrawn = "TransitionBottomTop" Then
            Call drawLeft(logicCommands(i), globalLineWidth, True, True)
        End If
    
        'Debug.Print logicCommands(i).commandDrawn
    Next i
End If
    
'Set column width
ActiveSheet.Cells.ColumnWidth = 8.53
For i = 0 To UBound(logicCommands)
    
    If InStr(logicCommands(i).commandDrawn, "Transition") > 0 Then
        If logicCommands(i).customWidth <> "" Then
            logicCommands(i).cell.ColumnWidth = 8.53 * (logicCommands(i).customWidth / 100)
        Else
            logicCommands(i).cell.ColumnWidth = 8.53 * (globalTransitionsCellWidth / 100)
        End If
    End If

    'Debug.Print logicCommands(i).cell.Address & " / " & logicCommands(i).customWidth

Next i
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
    
End Sub

Public Sub cellMarkType()

    cType = getCellMarkType(ActiveCell)
    MsgBox cType

End Sub

Function drawBottom(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)
    
    Set tgtCell = ci.cell.Offset(2, 0)
    
    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With

    If skipDrawState = False Then
        ci.commandDrawn = "Bottom"
    End If

End Function

Function drawTop(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)

    Set tgtCell = ci.cell.Offset(2, 0)

    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With
    
    'ActiveSheet.Cells.ColumnWidth = 8.53
    
    If skipDrawState = False Then
        ci.commandDrawn = "Top"
    End If

End Function

Function drawBus(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, text As String, Optional skipDrawState As Boolean)

    Set tgtCell = ci.cell.Offset(2, 0)

    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With
    With tgtCell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With
    
    tgtCell.Value2 = text
    
    If skipDrawState = False Then
        ci.commandDrawn = "Bus"
    End If

End Function

Function drawLeft(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)

    Set tgtCell = ci.cell.Offset(2, 0)

    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With

    If skipDrawState = False Then
        ci.commandDrawn = "Left"
    End If

End Function

Function drawRight(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)

    Set tgtCell = ci.cell.Offset(2, 0)

    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With

    If skipDrawState = False Then
        ci.commandDrawn = "Right"
    End If

End Function

Function drawBusTransition(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)

    Set tgtCell = ci.cell.Offset(2, 0)

    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlDiagonalDown)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With
    With tgtCell.Borders(xlDiagonalUp)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With
    
    If skipDrawState = False Then
        ci.commandDrawn = "BusTransition"
    End If
    
End Function

Function drawTransitionBottomTop(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)
    
    Set tgtCell = ci.cell.Offset(2, 0)
    
    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlDiagonalUp)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With

    If skipDrawState = False Then
        ci.commandDrawn = "TransitionBottomTop"
    End If

End Function

Function drawTransitionTopBottom(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)
    
    Set tgtCell = ci.cell.Offset(2, 0)
    
    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Borders(xlDiagonalDown)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .weight = getXLWeight(weight)
    End With

    If skipDrawState = False Then
        ci.commandDrawn = "TransitionTopBottom"
    End If

End Function

Function drawU(ByRef ci As commandItem, weight As Integer, preserveFormat As Boolean, Optional skipDrawState As Boolean)

    Set tgtCell = ci.cell.Offset(2, 0)

    If preserveFormat = False Then
        tgtCell.ClearFormats
    End If

    With tgtCell.Interior
        .Pattern = xlChecker
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    If skipDrawState = False Then
        ci.commandDrawn = "Unknown"
    End If
    
End Function

Function getCellMarkType(tgtCell As Range)

If tgtCell.Interior.Pattern = xlChecker Then
    getCellMarkType = "U"
End If

If tgtCell.Borders(xlDiagonalDown).LineStyle = xlContinuous And tgtCell.Borders(xlDiagonalUp).LineStyle = xlContinuous Then
    getCellMarkType = "X"
End If

If tgtCell.Borders(xlEdgeTop).LineStyle = xlContinuous And tgtCell.Borders(xlEdgeBottom).LineStyle = xlContinuous Then
    getCellMarkType = "="
ElseIf tgtCell.Borders(xlEdgeTop).LineStyle = xlContinuous Then
    getCellMarkType = "1"
ElseIf tgtCell.Borders(xlEdgeBottom).LineStyle = xlContinuous Then
    getCellMarkType = "0"
ElseIf tgtCell.Borders(xlDiagonalUp).LineStyle = xlContinuous Then
    getCellMarkType = "T1"
ElseIf tgtCell.Borders(xlDiagonalDown).LineStyle = xlContinuous Then
    getCellMarkType = "T0"
Else
    getCellMarkType = "n.a."
End If

End Function

Function markErrorCell(tgtCell As Range)

    With tgtCell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Function

Function getXLWeight(weight As Integer)
    Select Case weight
        Case 1
            getXLWeight = xlThin
        Case 2
            getXLWeight = xlMedium
        Case 3
            getXLWeight = xlThick
    End Select
End Function
