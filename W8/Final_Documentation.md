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

'I found that without a wait command, IE would sometimes not load still. Waiting 1 second seemed to fix this issue.
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
With a list of Game_ID's and their corresponding URL's, we could now begin capturing the data from Steam's store pages. This is the real "meat" of the process


