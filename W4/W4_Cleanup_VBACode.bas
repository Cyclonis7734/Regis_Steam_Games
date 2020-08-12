Attribute VB_Name = "W4_Cleanup_VBACode"
Option Explicit

Public Function CountItemsByDelim(strText As String, strDelim As String) As Integer
Dim intComCount As Integer
Dim intLet As Integer
Dim strLet As String

intComCount = 0
intLet = 0

If strText <> "" Then
    
    Do
        intLet = intLet + 1
        strLet = Mid(strText, intLet, 1)
        If strLet = strDelim Then
            intComCount = intComCount + 1
        End If
    
    Loop Until intLet >= Len(strText)
    
Else
    intComCount = -1
End If
    
CountItemsByDelim = intComCount + 1

End Function
