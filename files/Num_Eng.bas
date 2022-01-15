Attribute VB_Name = "Num_Eng"
Sub Num_Eng()
    
Dim cel As Range

For Each cel In Selection

    If Not IsEmpty(cel) And cel <> 0 Then
    
        If IsNumeric(CStr(cel)) Then
            If cel > 0 Then
            cel = NumToEng(CStr(cel))
            Else: cel = "-" & NumToEng(CStr(Abs(cel)))
            End If
        Else: cel = EngToNum(CStr(cel))
        
        End If
    End If
Next cel

End Sub


Function NumToEng(mNum As Double)

Dim num As Variant
Dim num2 As Double
num = mNum
Dim expArr As Variant
Dim letArr As Variant

expArr = Array(-21, -18, -15, -12, -9, -6, -3, 0, 3, 6, 9, 12, 15, 18, 21, 24, 27)
letArr = Array("y", "z", "a", "f", "p", "n", "u", "m", "", "K", "M", "G", "T", "P", "E", "Z", "Y")

Dim isOK As Boolean
Dim exp As Integer
Dim i As Integer

i = 0

Do Until isOK = True
exp = expArr(i)

If num < 10 ^ exp Then
    isOK = True
    num2 = num / 10 ^ (exp - 3)
    NumToEng = num2 & letArr(i)

End If

i = i + 1

If i = UBound(expArr) + 1 Then
    NumToEng = "Out of range"
    Exit Function
End If
Loop

End Function

Function EngToNum(mNum As String)

Dim mUnit As String
Dim mNum2 As Double
mUnit = Right(mNum, 1)
mNum2 = Left(mNum, Len(mNum) - 1)

Dim expArr As Variant
Dim letArr As Variant
Dim isOkay As Boolean
Dim i As Integer
Dim num As Double

i = 0

expArr = Array(-21, -18, -15, -12, -9, -6, -3, 0, 3, 6, 9, 12, 15, 18, 21, 24, 27)
letArr = Array("y", "z", "a", "f", "p", "n", "u", "m", "", "K", "M", "G", "T", "P", "E", "Z", "Y")

Do Until isOkay = True

If mUnit = letArr(i) Then
    
    num = mNum2 * 10 ^ (expArr(i) - 3)
    EngToNum = num
    isOkay = True
End If

i = i + 1

Loop

End Function

