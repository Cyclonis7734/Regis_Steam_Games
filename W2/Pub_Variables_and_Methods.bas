Attribute VB_Name = "Pub_Variables_and_Methods"
Option Explicit

Public Const strSteamMain = "https://store.steampowered.com/search/?sort_by=&sort_order=0&category1=998&filter=topsellers&page="

'enum separated by 10's, in the event that I need to add info between existing entries
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

'Does exactly what it says it does, by iterating through the text string
'given, and replacing any text that is equal to two spaces together.
'After that, it removes leading spaces that exist in the given string.
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

'Removes excessive carriage returns. First it replaces them with --vbLine--.
'This was done to ensure that the replacement was working, and allow for
'replacement of the various characters that resemble white space/carriage
'returns in one single grouping of replace statements. I ran into many issues
'during execution of the various cleanup functions I created, and this wound
'up making it very clear as to what was happening at each cleanup step. It would
'be easier to just set the replace function's replacement string to empty, but
'testing using the replacement with --vbLine-- was made much easier.
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

'Remove excess carriage returns from the beginning of the text string.
booDone = False
Do
    If Left(strText, 10) = "--vbLine--" Then
        strText = Mid(strText, 11, 999999)
    Else
        booDone = True
    End If
Loop Until booDone
'Debug.Print "----Replace leading CR's: " & strText

'Actually replace any remaining carriage returns with a single
'type of carriage return (vbNewLine).
strText = Replace(strText, "--vbLine--", vbNewLine)

RemoveExcessCarriageReturns = strText
End Function

'Loop through the given string, removing anything that's not a number.
'Additionally, a set of characters could be passed, which could be added
'to the string of characters to keep.
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

'This function was, usually, only called after the RemoveExcessCarriageReturns function call
Public Function RemoveAllCRs(strText As String) As String
RemoveAllCRs = Replace(strText, vbNewLine, "")
End Function

'Test sub
Public Sub testRemovals()
Dim strPass As String
strPass = "   " & vbNewLine & "fasdfsadf sa  sadfsadf sdaf saf sf asf  " & vbNewLine & vbNewLine & vbNewLine & vbNewLine & " fafsd fs asf saf afsadfsaf   fasfsdfsdafds gdsagdfgsd   gdf s"
strPass = RemoveExcessCarriageReturns(RemoveExcessSpaces(strPass))
Debug.Print strPass
End Sub

'Create a string by parsing a passed text string, split by any
'carriage returns. Then, get a specific value from that array.
Public Function ArrayByCR(strText As String, intIndex As Integer) As String
Dim arrVals As Variant

arrVals = Split(strText, vbNewLine)
'return nothing, in the event that the given index for the array is not available.
'This is achieved by ignoring errors during the index selection.
On Error Resume Next
ArrayByCR = arrVals(intIndex)
On Error GoTo 0

End Function







'Place holder to stop scrolling issue
