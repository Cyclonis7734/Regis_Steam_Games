Attribute VB_Name = "Cleanup_ESRBWhy"
Option Explicit

Public Sub MoveESRBWhys(longR As Long)
Dim wsM As Worksheet
Dim intC As Integer

Set wsM = ThisWorkbook.Worksheets("ESRBWhy Breakout")

With wsM
    For intC = 2 To 26
        .Range(.Cells(2, intC), .Cells(longR, intC)).Copy
        .Cells((longR * (intC - 1)) + 1, 1).PasteSpecial xlPasteValues
    Next intC

End With


End Sub

Public Sub RunSubs()
'1162 = number of rows with unique values when removing duplicates
'263 = number of rows after initial cleanup
MoveESRBWhys 263
End Sub

Public Function CleanChars(strVals As String) As String
Dim strFin As String
Dim longLetter As Long
Dim strLetter As String

strFin = ""

'Clear out anything that's not alphanumeric, a comma, period, ampersand, colon, or semicolon and
'replace them with spaces. TRIM function in Excel will remove excess spaces between words/letters.
'The PROPER function will set all beginning letters of perceived words as capitalized.
For longLetter = 1 To Len(strVals)
    strLetter = Mid(strVals, longLetter, 1)
    If InStr(1, " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789,.&:;", strLetter, vbTextCompare) > 0 Then
        strFin = strFin & strLetter
    Else
        strFin = strFin & " "
    End If
Next longLetter

'Remove leading spaces
Do
strLetter = Mid(strFin, 1, 1)
If strLetter = " " Then
    strFin = Mid(strFin, 2, 9999)
End If
Loop Until strLetter <> " "

'Remove ending spaces and/or special characters
Do
strLetter = Right(strFin, 1)
If strLetter = " " Or _
   strLetter = "." Or _
   strLetter = "," Or _
   strLetter = ";" Then
    strFin = Left(strFin, Len(strFin) - 1)
Else
    strLetter = ""
End If

Loop Until strLetter = ""


CleanChars = strFin

End Function

Public Function CleanCharsSTRICT(strVals As String)
Dim strFin As String
Dim longLetter As Long
Dim strLetter As String

strFin = ""

'Clear out anything that's not alphanumeric
For longLetter = 1 To Len(strVals)
    strLetter = Mid(strVals, longLetter, 1)
    If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", strLetter, vbTextCompare) > 0 Then
        strFin = strFin & strLetter
    End If
Next longLetter

CleanCharsSTRICT = strFin
End Function










