Attribute VB_Name = "Main"
Option Explicit
'https://www.guru99.com/data-scraping-vba.html

'Public Const strSteamMain As String = "https://store.steampowered.com/"
'Public Const strSteamSearch As String = "search/?category1=998&filter=topsellers"
'Public Const strSteamGameTemplate As String = "app/"


Public Sub GetGamesListingFromSteam()
Dim ie As New InternetExplorer
Dim ieD As New HTMLDocument
Dim ecoll As Object
Dim ecoll2 As Object
Dim ecoll3 As Object
Dim elem As Object
Dim longPasteRow As Long
Dim strPasteRow As String
Dim wsD As Worksheet
Dim longCounter As Long
Dim booStop As Boolean
Dim intPage As Integer
Dim strPage As String
Dim intMaxPg As Integer
Dim strAppID As String
Dim strHref As String

longCounter = 0
longPasteRow = 1
strPasteRow = Trim(Str(longPasteRow))
Set wsD = ThisWorkbook.Worksheets("Data")
wsD.Cells.ClearContents
wsD.Range("A:A").NumberFormat = "@"
strPage = "1"


'-------------------------------------------------------------------------
'Handle all things IE opening and navigating to webpage
'-------------------------------------------------------------------------
'ie.Visible = True
'ie.Navigate strSteamMain & strSteamSearch
ie.Navigate strSteamMain & strPage

Do While ie.Busy
    DoEvents
Loop

Application.Wait Now + TimeValue("0:00:01")
Set ieD = ie.Document
Set ecoll = ieD.getElementsByTagName("a")
Set ecoll2 = ieD.getElementsByClassName("search_pagination_right")
Set ecoll3 = ecoll2.Item(, 1).Children(2)
intMaxPg = CInt(ecoll3.textContent)

'Debug.Print intMaxPg
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


For intPage = 1 To intMaxPg
    strPage = Trim(Str(intPage))
'    ie.Quit
'    Set ie = Nothing
'    Set ie = New InternetExplorer
    ie.Navigate strSteamMain & strPage
    
    Do While ie.Busy
        DoEvents
    Loop
    Application.Wait Now + TimeValue("0:00:01")
    
    Set ieD = ie.Document
    Set ecoll = ieD.getElementsByTagName("a")
    
    For Each elem In ecoll
'        longCounter = longCounter + 1
'        Debug.Print elem.getAttribute("data-ds-appid")
        strAppID = ""
        On Error Resume Next
            strAppID = elem.getAttribute("data-ds-appid")
            strHref = elem.href
        On Error GoTo 0
        'Debug.Print strHref
        'If InStr(1, strHref, "/app/", vbTextCompare) > 0 Then
        If strAppID <> "" Then
            wsD.Range("A" & strPasteRow).Value = strAppID
            wsD.Range("B" & strPasteRow).Value = strHref
            longPasteRow = longPasteRow + 1
            strPasteRow = Trim(Str(longPasteRow))
        End If
    Next elem

Next intPage

ie.Quit
Set ie = Nothing
Set ecoll = Nothing
Set ecoll2 = Nothing
Set ecoll3 = Nothing

End Sub




'__________________________________________________________________________________________________________________
'__________________________________________________________________________________________________________________
Public Sub GetGameDetails(strWebPage As String, ws As Worksheet, longRow As Long)
Dim ie As New InternetExplorer
Dim ieD As New HTMLDocument
Dim intLoo As Integer
Dim ecoll As Object
Dim strVal As String
Dim strR As String
Dim objViewBtn As Object
Dim intDBNum As Integer
Dim booDiscount As Boolean

strR = Trim(Str(longRow))
ie.Navigate strWebPage
booDiscount = False
'ie.Visible = True

Do While ie.Busy
    DoEvents
Loop

Application.Wait Now + TimeValue("0:00:01")

Set ieD = ie.Document

Application.Wait Now + TimeValue("0:00:01")

If InStr(1, ie.LocationURL, "agecheck/") > 0 Then
    Set objViewBtn = ieD.getElementsByClassName("btnv6_blue_hoverfade btn_medium")
    objViewBtn(0).Click
End If
    
'Do While ie.Busy
'    DoEvents
'Loop
Application.Wait Now + TimeValue("0:00:02")
'Set ecoll = ieD.getElementsByClassName("game_description_snippet")

'COL: C ---> Get game description
ws.Range("C" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("game_description_snippet").Item(0).textContent, ttDesc_Snippet)

On Error Resume Next
'COL: D & E ---> Get % users who Like game, and total reviews given, separated by a Pipe symbol
strVal = ReturnCorrectedText(ieD.getElementsByClassName("user_reviews_summary_row").Item(1).textContent, ttReviews_Text)
ws.Range("D" & strR).Value = Left(strVal, InStr(1, strVal, "|", vbTextCompare) - 1)
ws.Range("E" & strR).Value = Mid(strVal, InStr(1, strVal, "|", vbTextCompare) + 1, 99)

'COL: F ---> Get Release1
ws.Range("F" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("release_date").Item(0).textContent, ttRelease_Date)


'Application.Wait Now + TimeValue("0:00:01")
'COL: G & H ---> Get Developer & Publisher
ws.Range("G" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("dev_row").Item(0).textContent, ttDeveloper)
ws.Range("H" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("dev_row").Item(1).textContent, ttPublisher)

'COL: I & J ---> Get ESRB Rating & ESRB Reason for Rating
ws.Range("I" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("game_rating_icon").Item(0).innerHTML, ttESRB)
ws.Range("J" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("descriptorText").Item(0).outerText, ttESRBWhy)
On Error GoTo 0

Application.Wait Now + TimeValue("0:00:01")

'ecoll = ieD.getElementsByClassName("details_block").Item

For intLoo = 0 To ieD.getElementsByClassName("details_block").Length
    If InStr(1, ieD.getElementsByClassName("details_block").Item(intLoo).outerText, "Genre: ", vbTextCompare) > 0 Or _
       InStr(1, ieD.getElementsByClassName("details_block").Item(intLoo).outerText, "Title: ", vbTextCompare) > 0 Then
        intDBNum = intLoo
        Exit For
    End If
Next intLoo

'COL: K:N ---> Get Title, Genre, Franchise, and Release2 columns

'For games that I have in my personal library, the page is structured differently.
'The below code caters to this fact.
'If InStr(1, ieD.getElementsByClassName("details_block").Item(0).outerText, "on record", vbTextCompare) > 0 Then
'    intDBNum = 4
'Else
'    intDBNum = 0
'End If

strVal = ieD.getElementsByClassName("details_block").Item(intDBNum).outerText
ws.Range("K" & strR).Value = ReturnCorrectedText(strVal, ttTitle)

strVal = ieD.getElementsByClassName("details_block").Item(intDBNum).outerText
ws.Range("L" & strR).Value = ReturnCorrectedText(strVal, ttGenre)

strVal = ieD.getElementsByClassName("details_block").Item(intDBNum).outerText
ws.Range("M" & strR).Value = ReturnCorrectedText(strVal, ttFranchise)

strVal = ieD.getElementsByClassName("details_block").Item(intDBNum).outerText
ws.Range("N" & strR).Value = ReturnCorrectedText(strVal, ttRelease_Date_2)

'COL: O ---> Get the Metacritic Score
'Sometimes there is no Metacritic score (usu because of bundled games into a single purchase item)
'In the event that this is the case, first show N/A, then assign the Metacritic score, if present
ws.Range("O" & strR).Value = "N/A"
On Error Resume Next
ws.Range("O" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("score high").Item(0).textContent, ttMetaCritic)
On Error GoTo 0

'Set ecoll = ieD.getElementsByClassName("game_purchase_price price")

'COL: P ---> Get price of first item on game page. This is usually the main game with no special purchases like DLC, etc.
On Error Resume Next
Set ecoll = ieD.getElementsByClassName("game_purchase_price price").Item(0)
If ecoll Is Nothing Then
    booDiscount = True
    Set ecoll = ieD.getElementsByClassName("discount_original_price").Item(0)
End If

If InStr(1, ecoll.textContent, "Demo", vbTextCompare) > 0 Or _
   InStr(1, ecoll.textContent, "Trial", vbTextCompare) > 0 Then
    If booDiscount Then
        Set ecoll = ieD.getElementsByClassName("discount_original_price").Item(1)
    Else
        Set ecoll = ieD.getElementsByClassName("game_purchase_price price").Item(1)
    End If
End If

ws.Range("P" & strR).Value = ReturnCorrectedText(ecoll.textContent, ttPrice)
On Error GoTo 0

'COL: Q ---> Get count of items available for purchase, including regular/main game option (should always be >= 1)
If booDiscount Then
    ws.Range("Q" & strR).Value = ieD.getElementsByClassName("discount_original_price").Length
Else
    ws.Range("Q" & strR).Value = ieD.getElementsByClassName("game_purchase_price price").Length
End If

'ws.Range("R" & strR).Value = ieD.getElementsByClassName("").Item(0).textContent
'ws.Range("S" & strR).Value = ieD.getElementsByClassName("").Item(0).textContent
'ws.Range("T" & strR).Value = ieD.getElementsByClassName("").Item(0).textContent
'ws.Range("U" & strR).Value = ieD.getElementsByClassName("").Item(0).textContent
'ws.Range("V" & strR).Value = ieD.getElementsByClassName("").Item(0).textContent
'ws.Range("W" & strR).Value = ieD.getElementsByClassName("").Item(0).textContent

'Debug.Print ie.LocationURL
ie.Quit
Set ie = Nothing
Set ieD = Nothing
Set ecoll = Nothing
strVal = ""
booDiscount = False

End Sub



Public Sub TestGameDeets()
GetGameDetails "https://store.steampowered.com/app/782330/DOOM_Eternal/", ThisWorkbook.Worksheets("Data"), 1
End Sub

Public Sub GameLooper()
Dim ws As Worksheet
Dim longR As Long
Dim strGameID As String
Dim intCounter As Integer
Dim longStopAt As Long

Set ws = ThisWorkbook.Worksheets("Data")
ws.Range("AA1").Value = "=COUNTA(A:A)"
intCounter = 0
longStopAt = ws.Range("AD1").Value

For longR = 2 To ws.Range("AA1").Value
    'if ending time for capturing is NOT available, then capture
    If ws.Range("Z" & Trim(Str(longR))).Value = "" Then
        
        'Increment counter, place Start time in Column Y, get Game ID #
        intCounter = intCounter + 1
        ws.Range("Y" & Trim(Str(longR))).Value = Time
        strGameID = ws.Range("A" & Trim(Str(longR))).Value
        
        'If multiple game id's given, assume 1st one is main game
        'and force it to be what is opened.
        If InStr(1, strGameID, ",", vbTextCompare) > 0 Then
            strGameID = Left(strGameID, InStr(1, strGameID, ",", vbTextCompare))
        End If
        
        'Get game info. When finished, place End Time in Column Z
        GetGameDetails "https://store.steampowered.com/app/" & strGameID, ws, longR
        ws.Range("Z" & Trim(Str(longR))).Value = Time
    End If
    
    
    'Allow pausing for ease of stopping code during mid running
    If intCounter > 50 Then
        intCounter = 0
        Application.StatusBar = "On game #: " & Trim(Str(longR + 1)) & " ---> Waiting..."
        Application.Wait Now + TimeValue("0:00:03")
    Else
        If longR >= longStopAt Then MsgBox "Requested stopping point reached.", vbOKOnly + vbInformation, "Stop Point Reached..."
        Application.StatusBar = "On game #: " & Trim(Str(longR + 1))
    End If
    
Next longR

End Sub


Public Function ReturnCorrectedText(strText As String, enumTextType As TextTypes) As String

'Based on given text type, decide how to format.
'First two select case options are for ESRB, as they
'did NOT require use of the functions nested in the
'RESRECR function call.
Select Case enumTextType
    Case TextTypes.ttESRB
        strText = Mid(strText, InStr(1, strText, "ESRB/", vbTextCompare) + 5, 999)
        strText = Left(strText, InStr(1, strText, ".", vbTextCompare) - 1)
        
    Case TextTypes.ttESRBWhy
        strText = RemoveExcessCarriageReturns(strText)
        strText = Replace(strText, vbNewLine, "|")

    Case Else
        strText = RESRECR(strText)
        
        Select Case enumTextType
            Case TextTypes.ttDesc_Snippet
                'Only needs the RESRECR Function call
                
            Case TextTypes.ttReviews_Text
                strText = RemoveNonNumeric(Replace(Mid(strText, InStr(1, strText, "- ", vbTextCompare) + 2, 20), "%", "|"), "|")
                
            Case TextTypes.ttRelease_Date
                strText = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 3, 12)
                
            Case TextTypes.ttDeveloper
                strText = RemoveAllCRs(Mid(strText, 13, 999))
                
            Case TextTypes.ttPublisher
                strText = RemoveAllCRs(Mid(strText, 13, 999))
                
            Case TextTypes.ttTitle
                strText = ArrayByCR(strText, 0)
                strText = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 2, 999)
            
            Case TextTypes.ttGenre
                strText = ArrayByCR(strText, 1)
                strText = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 2, 999)
            
            Case TextTypes.ttFranchise
                'Sometimes a Franchise listing is not available. If it is, capture, else leave blank
                If InStr(1, strText, "Franchise", vbTextCompare) > 0 Then
                    strText = ArrayByCR(strText, 4)
                    strText = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 2, 999)
                Else
                    strText = ""
                End If
            
            Case TextTypes.ttRelease_Date_2
                'Position of element in details_block class changes if Franchise is included or not.
                If InStr(1, strText, "Franchise", vbTextCompare) > 0 Then
                    strText = ArrayByCR(strText, 5)
                Else
                    strText = ArrayByCR(strText, 4)
                End If
                strText = Mid(strText, InStr(1, strText, ":", vbTextCompare) + 2, 999)
                            
            Case TextTypes.ttMetaCritic
                strText = RemoveAllCRs(strText)
            
            Case TextTypes.ttPrice
                strText = RemoveAllCRs(strText)
                
        End Select 'End of Other data
        
End Select 'End of Separate ESRB data


ReturnCorrectedText = strText

End Function


Public Function RESRECR(strText As String) As String
    strText = RemoveExcessSpaces(strText)
    strText = RemoveExcessCarriageReturns(strText)
    RESRECR = strText
End Function







'Place holder to stop scrolling issues

