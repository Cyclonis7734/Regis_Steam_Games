# Steam Games Data Analysis
Name: Thomas Seggerman
Course: MSDS 696 Practicum II
Date: 8/11/2020

## Project Details
This project was done for a college Practicum course at Regis University. The project's work was completed by using VBA for web-scraping from Steam's games listings on their website, then using Python for the remainder of the work. The course was 8 week's long, and the resulting list of games had over 16,000 titles.

## Initial Project Questions
Before beginning any Data Science project, it's good practice to determine what questions you are trying to answer. I have always wondered how Steam comes up with their pricing, but have never really investigated it myself. Steam often has games on sale, and the phrase "Steam Sale" has become synonymous with spending lots of money, for PC gamers. Steam also has a large availability of ratings from its user base, begging for questions on trending of reviews. That said, the main question that seem fair to ask at this point is:
<p><b>Can steam's store pages for games give us any insight into expectations for pricing or reviews of games?</b>

## Web-Scraping Fun
I have only done web-scraping a few times in my life, and it has never been pleasant. The amount of issues that occur when attempting to gather information from template web pages are directly related to how "clean" a website's data is kept. Needless to say, Steam, with it's thousands of games, is not without its issues in its game database! In deciding how to approach gathering this data, some important distinctions were made:
1. We would focus on games, and NOT software sold by Steam.
2. Narrow down the available games list by filtering for the Top Sellers listings.
3. Games that were in a bundle, will ONLY have the first title listed, as the focus of the group.

Below, we'll take a look at the various steps that were taken to obtain the data from Steam's store pages for its games. Up first is the code that was dedicated to creating the initial list of games. The code first creates an InternetExplorer object, then navigates to the Top Sellers page on Steam, opening up the appropriate location where the list of games will populate on a specific number of pages. A variable is used to hold the max number of pages, and then a loop begins to iterate through all items in the collection of game-list-objects on the web page. If each item in the list qualifies itself as a game, its Game_ID is captured, along with the URL for the game's store page.

```vba
Public Sub GetGamesListingFromSteam()
'Declare Variables
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

'Set a few of the variables
longCounter = 0
longPasteRow = 1
strPasteRow = Trim(Str(longPasteRow))
Set wsD = ThisWorkbook.Worksheets("Data")
wsD.Cells.ClearContents
wsD.Range("A:A").NumberFormat = "@"   'set the format of column A to text
strPage = "1"


'-------------------------------------------------------------------------
'Handle all things IE opening and navigating to webpage
'-------------------------------------------------------------------------
'ie.Visible = True
'ie.Navigate strSteamMain & strSteamSearch
ie.Navigate strSteamMain & strPage

'Loop aimlessly while InternetExplorer may be loading the page you've navigated to
Do While ie.Busy
    DoEvents
Loop
Application.Wait Now + TimeValue("0:00:01")

'Set various IE objects for use in discovering information about the web pages.
Set ieD = ie.Document
Set ecoll = ieD.getElementsByTagName("a")
Set ecoll2 = ieD.getElementsByClassName("search_pagination_right")
Set ecoll3 = ecoll2.Item(, 1).Children(2)
intMaxPg = CInt(ecoll3.textContent)

'Debug.Print intMaxPg
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


'Begin looping through each page
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
    
    'Loop through the game-list-objects collection
    'if /app/ is found in the URL and the Game_ID is NOT blank, then capture the object's data
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

'Clear out the memory of these objects in the event of an improper
'memory disposal issue in VBA.
ie.Quit
Set ie = Nothing
Set ecoll = Nothing
Set ecoll2 = Nothing
Set ecoll3 = Nothing

End Sub
```

### The Massive Undertaking
With a list of Game_ID's and their corresponding URL's, we could now begin capturing the data from Steam's store pages. This is the real "meat" of the process, and contains multiple steps which are designed to capture various points of data from each game's store page. Just FYI, there are a few functions called in the below code, that are available in a different module, shown later down below.

```vba
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

'Handles issue where an age check page loads, to ensure that the user is at least
'18 years old. It was easier to login and have Steam save my age, rather than
'trying to find the objects that hold the Month, Day, and Year values, and then
'update them. After logging in, the code would just be to hit the View Page button
'on the page. Then you are brought to the expected store page.
If InStr(1, ie.LocationURL, "agecheck/") > 0 Then
    Set objViewBtn = ieD.getElementsByClassName("btnv6_blue_hoverfade btn_medium")
    objViewBtn(0).Click
End If
Application.Wait Now + TimeValue("0:00:02")

'NOTE: The function "ReturnCorrectedText" is used to handle formatting and cleanup of the various
'      data that is being gathered. The steps to clean it up are in that function, rather than
'      here in the IE scraping code.
'COL: C ---> Get game description
ws.Range("C" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("game_description_snippet").Item(0).textContent, ttDesc_Snippet)

On Error Resume Next
'COL: D & E ---> Get % users who Like game, and total reviews given, separated by a Pipe symbol
strVal = ReturnCorrectedText(ieD.getElementsByClassName("user_reviews_summary_row").Item(1).textContent, ttReviews_Text)
ws.Range("D" & strR).Value = Left(strVal, InStr(1, strVal, "|", vbTextCompare) - 1)
ws.Range("E" & strR).Value = Mid(strVal, InStr(1, strVal, "|", vbTextCompare) + 1, 99)

'COL: F ---> Get Release1
ws.Range("F" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("release_date").Item(0).textContent, ttRelease_Date)

'COL: G & H ---> Get Developer & Publisher
ws.Range("G" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("dev_row").Item(0).textContent, ttDeveloper)
ws.Range("H" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("dev_row").Item(1).textContent, ttPublisher)

'COL: I & J ---> Get ESRB Rating & ESRB Reason for Rating
ws.Range("I" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("game_rating_icon").Item(0).innerHTML, ttESRB)
ws.Range("J" & strR).Value = ReturnCorrectedText(ieD.getElementsByClassName("descriptorText").Item(0).outerText, ttESRBWhy)
On Error GoTo 0

Application.Wait Now + TimeValue("0:00:01")

For intLoo = 0 To ieD.getElementsByClassName("details_block").Length
    If InStr(1, ieD.getElementsByClassName("details_block").Item(intLoo).outerText, "Genre: ", vbTextCompare) > 0 Or _
       InStr(1, ieD.getElementsByClassName("details_block").Item(intLoo).outerText, "Title: ", vbTextCompare) > 0 Then
        intDBNum = intLoo
        Exit For
    End If
Next intLoo

'COL: K:N ---> Get Title, Genre, Franchise, and Release2 columns
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

'Determine if game is a Demo or Trial. If so, handle price obtaining differently.
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


'The actual sub that begins and ends the looping process.
Public Sub GameLooper()
Dim ws As Worksheet
Dim longR As Long
Dim strGameID As String
Dim intCounter As Integer
Dim longStopAt As Long

Set ws = ThisWorkbook.Worksheets("Data")
ws.Range("AA1").Value = "=COUNTA(A:A)"
intCounter = 0

'This variable was to have a custom stopping point, rather than just the
'end of the file. The total time to capture ALL the game's data was over
'20 hours. This was necessary to not require manual stopping of the code
'while running, using the control+pause|break keyboard command.
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


'This is the Function that handles formatting and cleanup of the various text
'being captured from each game's store page.
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

'RESRECR = abbreviations of the words for the functions
'being called by this function. These 2 functions were
'called so often, that it made sense to reduce them to
'a single function, and cut the lines of code down.
Public Function RESRECR(strText As String) As String
    strText = RemoveExcessSpaces(strText)
    strText = RemoveExcessCarriageReturns(strText)
    RESRECR = strText
End Function
```

The next bits of code, contain some functions that were designed to handle various cleanup steps. The module they are on is for hosting public variables and methods which can be used by the entirety of the Excel file hosting this VBA code.

```vba
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
```

