Attribute VB_Name = "LetterButton"
Option Explicit
Private buttonSheet As Worksheet
'Variables for working with Word and creating letter
Private WordApp As Word.Application
Private WordDoc As Word.Document
'Advisor names taken from button sheet
Private adv1First As String
Private adv1Last As String
Private adv2First As String
Private adv2Last As String
Sub BuildLetter(Optional bothClicked As Boolean)
    LogData.TimeStart
    'TODO: if there's an RMD and it's the months we add RMD stuff, have a message box saying to add something to the letter; don't print the letter
    'Turn off screen updating
    BothButton.UpdateScreen "Off"
    
    'Error handling
    'TODO: Make this better
    On Error GoTo MacroBroke
    
    'Set global variables
    If Not bothClicked Then
        BothButton.SetGlobals
    End If
    SetLetterGlobals
    
    'Log macro use - Don't log if this is the test builder
    If ThisWorkbook.Name <> "Test Report Builder.xlsm" Then
        LogData.WriteLog "Letter", BothButton.HouseholdName, True
    End If
    
    'Open the letter
    OpenLetter
    
    'Add client names and advisor names to letter
    ProcessLetter
    
    'Print the letter
    PrintLetter
    
    'Save the letter
    SaveLetter
    
    'Reset the global variables
    ResetLetterGlobals True
    If Not bothClicked Then
        'Just the letter was built; reset the global variables in BothButton
        BothButton.ResetGlobals
    End If
    
    'Turn the screen updating back on
    BothButton.UpdateScreen "On"
    
    'Log the time
    LogData.TimeEnd
    If Not bothClicked Then
        LogData.LogTime
    End If
    Exit Sub
    
MacroBroke:
    'TODO: Make the error handling better
    ResetLetterGlobals False
    ErrorHandling.ErrorAndStop
End Sub
Sub SetLetterGlobals()
    'Set the button's workbook and worksheet
    Set buttonSheet = BothButton.ReportBuilderSheet
    
    'Find who's signing the letter
    GetAdvisors
End Sub
Sub GetAdvisors()
'Find which advisor checkboxes are checked
If buttonSheet.Shapes("cbxDan").OLEFormat.Object.Object.Value = True Then
    'If Dan is checked, he will always be the first advisor
    adv1First = "Dan"
End If
If buttonSheet.Shapes("cbxRyan").OLEFormat.Object.Object.Value = True Then
    'Ryan's the first advisor if Dan isn't; otherwise he's the second advisor
    If adv1First = vbNullString Then
        adv1First = "Ryan"
    Else
        adv2First = "Ryan"
    End If
End If
If buttonSheet.Shapes("cbxRachel").OLEFormat.Object.Object.Value = True Then
    'If Rachel is checked, she will always be the second advisor
    adv2First = "Rachel"
End If

'Get the advisors' last names from their first names
adv1Last = GetLast(adv1First)
adv2Last = GetLast(adv2First)

'Check for errors
If adv1First = vbNullString Then
    'None of the advisors were checked
    ErrorHandling.ErrorAndStop "Please select at least one advisor to sign the letter", "Someone's gotta sign this thing"
End If
    
If adv1Last = vbNullString Or (adv2First <> vbNullString And adv2Last = vbNullString) Then
    'At least one of the last names are missing
    ErrorHandling.ErrorAndContinue "Something went wrong and at least one of the advisors' last names are missing. It will need to be added to the letter manually.", "This was Steve's fault"
End If
End Sub
Function GetLast(advFirst) As String
    If advFirst = "Dan" Then
        GetLast = "Budinger"
    ElseIf advFirst = "Ryan" Then
        GetLast = "Wempe"
    ElseIf advFirst = "Rachel" Then
        GetLast = "Brown"
    Else
        GetLast = vbNullString
    End If
End Function
Sub OpenLetter()
    'Get the letter's location
    Dim letterPath As String
    letterPath = LetterFileLocation
 
    'Open the file, but keep Word hidden until it's done processing
    Set WordApp = New Word.Application
    WordApp.Visible = False
    Set WordDoc = WordApp.Documents.Open(letterPath, ReadOnly:=True)
    WordDoc.Activate
    WordDoc.ActiveWindow.View.ReadingLayout = False
End Sub
Function LetterFileLocation() As String
    'Build the file name using a directory, folder, file, and extension
    'Use provided file if available. If not, use the default file. If that's not available, open a file dialog
    
    'Get the file
    Dim file As String
    file = buttonSheet.Range("FileName").Value
    If file = "" Then
        file = DefFile
    End If
    
    'Get the folder
    Const defFolder As String = "FPIS - Operations\Client Communications\FPIS - Inv Rec Letters\"
    Dim folder As String
    folder = buttonSheet.Range("FileFolder").Value
    If folder = "" Then
        folder = defFolder
    Else
        'Ensure the folder ends with a \
        If Right(folder, 1) <> "\" Then
            folder = folder & "\"
        End If
    End If
    
    'Get the extension
    Const defExtension As String = ".docx"
    Dim extension As String
    extension = buttonSheet.Range("FileExtension").Value
    If extension = "" Then
        extension = defExtension
    End If
    
    'Put the file name together
    Const defDir As String = "Z:\"
    Dim letterPath As String
    letterPath = defDir & folder & file & extension
    
    'If the provided file doesn't exist, use the default file
    If Dir(letterPath) = "" Then
        letterPath = defDir & defFolder & DefFile & defExtension
        
        'If the default location doesn't exist, select a file to open
        If Dir(letterPath) = "" Then
            letterPath = FindPath
        End If
    End If
    
    'Return the file location
    LetterFileLocation = letterPath
End Function
Function DefFile() As String
    'Find which quarter it is and the corresponding months for the default file
    Dim quarter As String
    Dim quarterMonths As String
    If Month(Date) >= 1 And Month(Date) <= 3 Then
        quarter = "1st"
        quarterMonths = "Jan, Feb, Mar"
    ElseIf Month(Date) >= 4 And Month(Date) <= 6 Then
        quarter = "2nd"
        quarterMonths = "Apr, May, Jun"
    ElseIf Month(Date) >= 7 And Month(Date) <= 9 Then
        quarter = "3rd"
        quarterMonths = "Jul, Aug, Sep"
    Else
        quarter = "4th"
        quarterMonths = "Oct, Nov, Dec"
    End If
    DefFile = quarter & " Quarter " & Year(Date) & " - " & quarterMonths & " TRX"
End Function
Function FindPath() As String
    'Set the initial folder
    Dim initialFolder As String
    initialFolder = "Z:\FPIS - Operations\Client Communications\FPIS - Inv Rec Letters\"
    
    'If this folder isn't available, go to the top of the server
    If Dir(initialFolder) = "" Then
        initialFolder = "Z:\"
    End If
    
    'Show a file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Word Files", "*.docx; *.doc", 1
        .InitialFileName = initialFolder
        .Show
        
        'Get the selected file
        Dim fileSelected As Variant
        For Each fileSelected In .SelectedItems
            FindPath = .SelectedItems.Item(1)
        Next fileSelected
    End With
    
    'Quit the procedure if the user didn't select the type of file we need.
    If InStr(FindPath, ".doc") = 0 Then
        ErrorHandling.ErrorAndStop "No letter template was provided. Macro has halted.", "Can't have a letter without a letter"
    End If
End Function
Sub ProcessLetter()
    'Set the words to look for
    Dim targetWords() As Variant
    targetWords = Array("DATE", "Dear", "BUYSELLTARGET", "RMDLOC", "INSERT", "Regards")
    Dim targetWordsFound() As Boolean
    ReDim targetWordsFound(0 To UBound(targetWords))
    
    'Iterate through letter by paragraph to find target words
    Dim para As Integer
    Dim currentPara As String
    para = 1
    Do While para < WordDoc.Paragraphs.count And Not targetWordsFound(UBound(targetWords))
        'Get the paragraph's text
        currentPara = Trim(WordDoc.Paragraphs(para).Range.Text)
        
        'Loop through each target word to see if it's in the paragraph
        Dim targetWord As Integer
        For targetWord = 0 To UBound(targetWords)
            If InStr(currentPara, targetWords(targetWord)) > 0 And Not targetWordsFound(targetWord) Then
                'The paragraph contains the target word and that word hasn't been found yet. Replace it with the appropriate text
                targetWordsFound(targetWord) = ReplaceKeywords(para, targetWords(targetWord))
            End If
        Next targetWord
        
        'Go to the next paragraph
        para = para + 1
    Loop
End Sub
Function ReplaceKeywords(para As Integer, keyword As Variant) As Boolean
    'Iterate through words in the paragraph
    Dim keywordFound As Boolean
    Dim wrd As Integer
    wrd = 1
    Do While wrd < WordDoc.Paragraphs(para).Range.Words.count And keywordFound = False
        'Get the next word from the paragraph
        Dim paragraphRange As Word.Range
        Set paragraphRange = WordDoc.Paragraphs(para).Range
        Dim currentWordRange As Word.Range
        Set currentWordRange = paragraphRange.Words(wrd)
        Dim currentWord As String
        currentWord = currentWordRange.Text
        If currentWord = keyword Then
            'Call proper replace function
            If keyword = "DATE" Then
                currentWordRange = Replace(currentWord, keyword, ReplaceDate)
                keywordFound = True
            ElseIf keyword = "Dear" Then
                currentWordRange = Replace(currentWord, keyword, ReplaceDear)
                keywordFound = True
            ElseIf keyword = "BUYSELLTARGET" Then
                currentWordRange = Replace(currentWord, keyword, ReplaceTarget)
                keywordFound = True
            ElseIf keyword = "RMDLOC" Then
                currentWordRange = Replace(currentWord, keyword, ReplaceRMD)
                keywordFound = True
            ElseIf keyword = "INSERT" Then
                If ReplaceInsert = "" Then
                    currentWordRange = Replace(currentWord, keyword, ReplaceInsert)
                Else
                    currentWordRange = Replace(currentWord, keyword, vbCr & ReplaceInsert)
                End If

                'The disclaimers are after INSERT. Make sure the formatting doesn't change
                With WordDoc.Paragraphs(para).Range
                    WordDoc.Range(.Start, .End).Font.Size = 12
                    WordDoc.Range(.Start, .End).Font.Italic = False
                    WordDoc.Range(.Start, .End).Font.Bold = False
                End With
                keywordFound = True
            ElseIf keyword = "Regards" Then
                ReplaceRegards currentWordRange
                keywordFound = True
            End If
        End If
        
        'Go to the next word
        wrd = wrd + 1
    Loop
    
    ReplaceKeywords = keywordFound
End Function
Function ReplaceDate() As String
    'Get the current date, or next day's date if it's after 3:00
    Dim currentDate As String
    If Time < TimeValue("15:00:00") Then
        currentDate = MonthName(Month(Date), False) & " " & Day(Date) & ", " & Year(Date)
    Else
        If Weekday(Date, vbSunday) = 6 Then
            'If it's late on a Friday, date the letters for Monday
            currentDate = MonthName(Month(Date + 3), False) & " " & Day(Date + 3) & ", " & Year(Date + 3)
        Else
            'Late in the day and not Friday, date the letters for the next day
            currentDate = MonthName(Month(Date + 1), False) & " " & Day(Date + 1) & ", " & Year(Date + 1)
        End If
    End If
    ReplaceDate = currentDate
End Function
Function ReplaceDear() As String
    'Get client(s) first name(s)
    Dim fNames As String
    fNames = GetFirstNames
    ReplaceDear = "Dear " & fNames
End Function
Function GetFirstNames() As String
    'Get the household name
    Dim hhName As String
    hhName = BothButton.HouseholdName
    
    'Get the location of the ampersand, if it's there
    Dim ampPos As Integer
    ampPos = InStr(hhName, "&")
    
    'Parse the client(s) first name(s) from the household name
    Dim fName As String
    Dim spouseFirst As String
    If InStr(hhName, ",") > 0 Then
        'Household name has a comma, so it's "[last], [first] (& [spouse])". Find where the comma is and take everything after it
        Dim commaPos As Integer
        commaPos = InStr(hhName, ",")
        hhName = Trim(Right(hhName, Len(hhName) - commaPos))
        
        'Reset the ampersand position
        ampPos = InStr(hhName, "&")
        
        'Get first names
        If ampPos > 0 Then
            'Household name has spouse's name
            'Everything before the ampersand is the first name
            fName = Trim(Left(hhName, ampPos - 1))
            
            'Everything after the ampersand is the spouse's first name
            spouseFirst = Trim(Right(hhName, Len(hhName) - ampPos - 1))
            
            If InStr(spouseFirst, ",") > 0 Then
                'There's another comma, so the spouse has a different last name and it's "[last], [first] & [spouse last], [spouse first]".
                'Take everything after the second comma
                Dim spouseCommaPos As Integer
                spouseCommaPos = InStr(spouseFirst, ",")
                spouseFirst = Trim(Right(spouseFirst, Len(spouseFirst) - spouseCommaPos))
            End If
        Else
            'Household name is just one person, so the first name is what's left
            fName = hhName
        End If
        
        'Take out suffixes
        fName = Trim(Replace(fName, "Jr", ""))
        fName = Trim(Replace(fName, "Sr", ""))
    Else
        'Household name doesn't have a comma, so it's "[first] & [spouse] [last]" or "[first] [last] & [spouse] [spouse last]"
        'This is likely to give an incorrect name, but the error will bring this to user's attention
        Dim spacePos As Integer
        If ampPos > 0 Then
            'Household name has a spouse name. Assume it's "[first] & [spouse] [last]"
            'Everything before the ampersand is the first name
            fName = Trim(Left(hhName, ampPos - 1))
            
            'Everything between the ampersand and first space is (probably) the spouse's first name
            'Take out the ampersand and everything before it
            hhName = Trim(Right(hhName, Len(hhName) - ampPos - 1))
            
            'Everything before the space should be the spouse's first name
            spacePos = InStr(hhName, " ")
            spouseFirst = Trim(Left(hhName, spacePos - 1))
        Else
            'Single person
            'Everything before the space should be the first name
            spacePos = InStr(hhName, " ")
            fName = Trim(Left(hhName, spacePos - 1))
        End If
        
        'Throw error for incorrect format, but keep running macro
        Dim errorMessage As String
        errorMessage = "Household name is in incorrect format and names on letter may be incorrect. Proper format should be ""[last name], [first] & [spouse first]"""
        ErrorHandling.ErrorAndContinue errorMessage, "Look, I tried, ok?"
    End If
    
    GetFirstNames = CombineNames(fName, spouseFirst)
End Function
Function CombineNames(first As String, spouseFirst As String) As String
    If spouseFirst = "" Then
        'Single person
        CombineNames = first
    Else
        'Couple
        CombineNames = first & " and " & spouseFirst
    End If
End Function
Function ReplaceTarget() As String
    'Add the amount of equities bought or sold
    Dim buysell As Single
    buysell = AmountBoughtSold
    
    'If the buysell is negative, then we're selling equities
    Dim bstStr As String
    If buysell < 0 Then
        'Selling equities
        bstStr = "At this time, we will sell " & Format(-1 * buysell, "$#,##0") _
            & " of stock funds in order to move towards your " & EqTarget & " equity target"
    ElseIf buysell > 0 Then
        'Buying equities
        bstStr = "At this time, we will buy " & Format(buysell, "$#,##0") _
            & " of stock funds in order to move towards your " & EqTarget & " equity target"
    Else
        'At target
        bstStr = "At this time you are at your " & EqTarget & " equity target." _
            & "  We will move money around the investment style grid in order to buy" _
            & " low and sell high among different asset classes"
    End If
    
    ReplaceTarget = bstStr
End Function
Function AmountBoughtSold() As Double
    'Get the total bought or sold of FI or MMM
    Dim total As Double
    
    'Loop through each account in the list
    Dim acctList As Collection
    Set acctList = BothButton.AccountList
    Dim acct As Variant
    For Each acct In acctList
        'Look for these subclasses
        Dim fixedArray() As String
        fixedArray = Split("FI;MMM", ";")
        
        'Loop through each transaction in the account
        Dim trans As Variant
        For Each trans In acct.TradeList
            'Loop through each subclass that we're looking for
            Dim fixedSubclass As Variant
            For Each fixedSubclass In fixedArray
                If trans.subclass = fixedSubclass Then
                    'The transaction is fixed income or money market, add the trade value to the total
                    total = total + trans.Trade
                End If
            Next fixedSubclass
        Next trans
    Next acct
    
    'Since this is the total fixed income bought or sold, take the opposite of it to get the equities
    total = -1 * total
    
    'Round to the nearest 100
    total = Round(total / 100) * 100
    
    'Return the overall amount of equities bought or sold in the entire portfolio
    AmountBoughtSold = total
End Function
Function ReplaceRMD() As String
    If buttonSheet.Shapes("cbxRMD").OLEFormat.Object.Object.Value = True Then
        'There's at least one RMD. Put in the proper wording here
        'ReplaceRMD =
    Else
        'There's no RMD. Take out the placeholder
        ReplaceRMD = ""
    End If
End Function
Function ReplaceInsert() As String
    ReplaceInsert = buttonSheet.OLEObjects("lblInsert").Object.Caption
End Function
Sub ReplaceRegards(regardsRange As Word.Range) 'Add the advisor names, which appear after "Regards,"
    'Fill array with advisor names
    Dim advStr As String
    Dim advArray() As String
    advStr = "Dan#Budinger#Ryan#Wempe#Rachel#Brown"
    advArray = Split(advStr, "#")

    'Set a counter for the number of advisor names added
    Dim advNames As Integer
    advNames = 0
    
    With WordDoc
        'Loop through the words after regards until an advisor name is found
        Dim endRange As Word.Range
        Set endRange = WordDoc.Range(regardsRange.Start, .Range.End)
        
        Dim wrd As Integer
        wrd = 1
        Do While wrd < endRange.Words.count - 1 And advNames < 4
            Dim currentWord As String
            currentWord = Trim(endRange.Words(wrd).Text)
            
            'If the word is in the advisor array
            If currentWord <> "" And InStr(advStr, currentWord) > 0 Then
                Dim advIndex As Integer
                For advIndex = 0 To UBound(advArray)
                    If currentWord = advArray(advIndex) Then
                        'The word is in the advisor array
                        If advIndex Mod 2 = 0 Then
                            'The word is a first name
                            If advNames = 0 Then
                                'No advisors have been added, so add the first
                                endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), adv1First)
                            Else
                                'One advisor has been added, so add the second
                                endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), adv2First)
                            End If
                            advNames = advNames + 1
                        Else
                            'The word is a last name
                            If advNames = 1 Then
                                'No advisors have been added, so add the first
                                endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), adv1Last)
                            Else
                                'One advisor has been added, so add the second
                                endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), adv2Last)
                            End If
                            advNames = advNames + 1
                        End If
                    End If
                Next advIndex
            End If
            wrd = wrd + 1
        Loop
    End With
End Sub
Sub PrintLetter()
    'Make Word visible and bring it into focus (activate doesn't always do this; not sure why)
    WordApp.Visible = True
    WordApp.Activate
    
    If buttonSheet.Shapes("cbxRMD").OLEFormat.Object.Object.Value = False Then
        'No RMD, print the letter
        LogData.TimeEnd
        WordApp.PrintPreview = True
        Do While WordApp.PrintPreview = True
            DoEvents
        Loop
        LogData.TimeStart
    Else
        'There's an RMD, don't print the letter so the wording can be put in manually
    End If
End Sub
Sub SaveLetter()
    'Default file name is [Month] [Year].docx
    Dim saveFile As String
    saveFile = MonthName(Month(Date)) & " " & Year(Date)
    
    'Open save dialog with the default file name of the current month and current year
    LogData.TimeEnd
    Dim sfdSaveLetter As FileDialog
    Set sfdSaveLetter = WordApp.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveLetter.InitialFileName = BothButton.saveDir & saveFile
    sfdSaveLetter.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveLetter.SelectedItems.count > 0 Then
        savePath = sfdSaveLetter.SelectedItems.Item(1)
        WordDoc.SaveAs2 fileName:=savePath, ReadOnlyRecommended:=False
    End If
    LogData.TimeStart
End Sub
Sub ResetLetterGlobals(keepOpen As Boolean)
    Set buttonSheet = Nothing
    
    'Close Word document and application, and set them to nothing
    If Not keepOpen Then
        If Not WordDoc Is Nothing Then
            WordDoc.Close SaveChanges:=False
            Set WordDoc = Nothing
        End If
        If Not WordApp Is Nothing Then
            WordApp.Quit
            Set WordApp = Nothing
        End If
    End If
End Sub
