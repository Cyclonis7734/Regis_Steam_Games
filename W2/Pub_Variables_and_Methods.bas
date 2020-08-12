Attribute VB_Name = "Pub_Variables_and_Methods"
Option Explicit

Public Const strSteamMain = "https://store.steampowered.com/search/?sort_by=&sort_order=0&category1=998&filter=topsellers&page="

Public Enum TextTypes
    ttDesc_Snippet = 10
    ttReviews_Text = 20
    ttRelease_Date = 30
    ttDeveloper = 40
    ttPublisher = 50
    ttESRB = 60
    ttESRBWhy = 70
    ttTitle = 80
    ttGenre = 90
    ttFranchise = 100
    ttRelease_Date_2 = 110
    ttMetaCritic = 120
    ttPrice = 130
End Enum


Public Function RemoveExcessSpaces(strText) As String
Dim booDone As Boolean

booDone = False

Do
    If InStr(1, strText, "  ", vbTextCompare) = 0 Then booDone = True
    If Not booDone Then strText = Replace(strText, "  ", " ", 1)
    'Debug.Print strText
Loop Until booDone

booDone = False
Do
    If Left(strText, 1) = " " Then strText = Mid(strText, 2, 999999) Else booDone = True
Loop Until booDone

RemoveExcessSpaces = strText
End Function

Public Function RemoveExcessCarriageReturns(strText) As String
Dim booDone As Boolean


'Debug.Print "----Start: " & strText
booDone = False
strText = Replace(strText, vbCrLf, "--vbLine--")
strText = Replace(strText, Chr(10), "--vbLine--")
strText = Replace(strText, Chr(13), "--vbLine--")
strText = Replace(strText, Chr(9), "--vbLine--")

'Debug.Print "----Replace vbCrLf's: " & strText

Do
    If InStr(1, strText, "--vbLine--" & "--vbLine--", vbTextCompare) > 0 Then
        strText = Replace(strText, "--vbLine--" & "--vbLine--", "--vbLine--")
    Else
        booDone = True
    End If
Loop Until booDone

'Debug.Print "----Replace double CR's: " & strText

booDone = False
Do
    If Left(strText, 10) = "--vbLine--" Then
        strText = Mid(strText, 11, 999999)
    Else
        booDone = True
    End If
Loop Until booDone

'Debug.Print "----Replace leading CR's: " & strText

strText = Replace(strText, "--vbLine--", vbNewLine)

'       Left(strText, 1) = vbCrLf Or _
'       Left(strText, 1) = Chr(10) Or _
'       Left(strText, 1) = Chr(13)

RemoveExcessCarriageReturns = strText
End Function


Public Function RemoveNonNumeric(strText As String, Optional strListCharsKeep As String) As String
Dim longChar As Long
Dim strFin As String
Dim strLet As String

strFin = ""

For longChar = 1 To Len(strText)
    strLet = Mid(strText, longChar, 1)
    If strListCharsKeep = "" Then
        If InStr(1, "0123456789", strLet, vbTextCompare) > 0 Then strFin = strFin & strLet
    Else
        If InStr(1, "0123456789" & strListCharsKeep, strLet, vbTextCompare) > 0 Then strFin = strFin & strLet
    End If
Next longChar

RemoveNonNumeric = strFin

End Function

Public Function RemoveAllCRs(strText As String) As String
RemoveAllCRs = Replace(strText, vbNewLine, "")
End Function


Public Sub testRemovals()
Dim strPass As String
strPass = "   " & vbNewLine & "fasdfsadf sa  sadfsadf sdaf saf sf asf  " & vbNewLine & vbNewLine & vbNewLine & vbNewLine & " fafsd fs asf saf afsadfsaf   fasfsdfsdafds gdsagdfgsd   gdf s"
strPass = RemoveExcessCarriageReturns(RemoveExcessSpaces(strPass))
Debug.Print strPass
End Sub

Public Function ArrayByCR(strText As String, intIndex As Integer) As String
Dim arrVals As Variant

arrVals = Split(strText, vbNewLine)
On Error Resume Next
ArrayByCR = arrVals(intIndex)
On Error GoTo 0

End Function







'Place holder to stop scrolling issue
