VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_wordApp As Word.Application
Private m_wordDoc As Word.Document
Private m_household As clsHousehold
Private m_advisors() As String
Private m_insert As String

Public Property Get WordApp() As Word.Application
    Set WordApp = m_wordApp
End Property

Public Property Get WordDoc() As Word.Document
    Set WordDoc = m_wordDoc
End Property

Public Property Get household() As clsHousehold
    Set household = m_household
End Property

Public Property Get Advisors() As String()
    Advisors = m_advisors
End Property

Public Property Get Insert() As String
    Insert = m_insert
End Property

Public Sub ClassConstructor(household As clsHousehold)
    Set m_wordApp = New Word.Application
    m_wordApp.Visible = False
    Set m_wordDoc = OpenLetter
    Set m_household = household
End Sub

Private Function OpenLetter() As Word.Document
    'Get the letter's location
    Dim letterPath As String
    letterPath = LetterFileLocation
 
    'Open the file, but keep Word hidden until it's done processing
    Set OpenLetter = m_wordApp.Documents.Open(letterPath, ReadOnly:=True)
End Function

Private Function LetterFileLocation() As String
    'Start the file name at the top of the server directory
    Const defDir As String = "Z:\"
    Dim letterPath As String
    letterPath = defDir
    
    'Set the report builder's window
    Dim buttonWindow As New clsWindow
    Set buttonWindow = ClassBuilder.NewWindow("Report Builder")
    
    'Build the file name using a directory, folder, file, and extension
    'Use provided file in report builder if available. If not, use the default file. If that's not available, open a file dialog
    With buttonWindow.Book.Worksheets(1)
        'Get the folder
        Const defFolder As String = "FPIS - Operations\Client Communications\FPIS - Inv Rec Letters\"
        Dim folder As String
        folder = .Range("FileFolder").Value2
        If folder = "" Then
            folder = defFolder
        Else
            'Ensure the folder ends with a \
            If Right(folder, 1) <> "\" Then
                folder = folder & "\"
            End If
        End If
        
        'Add the folder to the path
        letterPath = letterPath & folder
        
        'Get the file
        Dim file As String
        file = .Range("FileName").Value2
        If file = "" Then
            file = DefFile
        End If
        
        'Add the file to the path
        letterPath = letterPath & file
        
        'Get the extension
        Const defExtension As String = ".docx"
        Dim extension As String
        extension = .Range("FileExtension").Value2
        If extension = "" Then
            extension = defExtension
        End If
        
        'Add the extension to the path
        letterPath = letterPath & extension
    End With
    
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

Private Function DefFile() As String
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
    
    'Return the default file name, e.g. "1st Quarter 2020 - Jan, Feb, Mar TRX"
    DefFile = quarter & " Quarter " & Year(Date) & " - " & quarterMonths & " TRX"
End Function

Private Function FindPath() As String
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

Public Sub Process()
    m_wordDoc.Activate
    m_wordDoc.ActiveWindow.View.ReadingLayout = False
    
    'Set the words to look for
    Dim targetWords() As String
    targetWords = Split("DATE,Dear,BUYSELLTARGET,RMDLOC,INSERT,Regards", ",")
    Dim targetWordsFound() As Boolean
    ReDim targetWordsFound(0 To UBound(targetWords))
    
    'Iterate through letter by paragraph to find target words
    Dim para As Integer
    Dim currentPara As String
    para = 1
    Do While para < m_wordDoc.Paragraphs.count And Not targetWordsFound(UBound(targetWords))
        'Get the paragraph's text
        currentPara = Trim(m_wordDoc.Paragraphs(para).Range.Text)
        
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

Private Function ReplaceKeywords(para As Integer, keyword As String) As Boolean
    'Set the report builder's window
    Dim buttonWindow As New clsWindow
    Set buttonWindow = ClassBuilder.NewWindow("Report Builder")
    
    'Find the keyword's location in the paragraph and get the text there
    Dim keywordRange As Word.Range
    Set keywordRange = FindKeyword(m_wordDoc.Paragraphs(para).Range, keyword)
    Dim keywordRangeText As String
    keywordRangeText = keywordRange.Text
    
    'Call proper replace function
    If keyword = "DATE" Then
        keywordRange = Replace(keywordRangeText, keyword, ReplaceDate)
    ElseIf keyword = "Dear" Then
        keywordRange = Replace(keywordRangeText, keyword, ReplaceDear)
    ElseIf keyword = "BUYSELLTARGET" Then
        keywordRange = Replace(keywordRangeText, keyword, ReplaceTarget)
    ElseIf keyword = "RMDLOC" Then
        Dim isRMDChecked As Boolean
        isRMDChecked = buttonWindow.Book.Worksheets(1).Shapes("cbxRMD").OLEFormat.Object.Object.value
        keywordRange = Replace(keywordRangeText, keyword, ReplaceRMD(isRMDChecked))
    ElseIf keyword = "INSERT" Then
        m_insert = buttonWindow.Book.Worksheets(1).OLEObjects("lblInsert").Object.Caption
        If ReplaceInsert = "" Then
            keywordRange = Replace(keywordRangeText, keyword, ReplaceInsert)
        Else
            keywordRange = Replace(keywordRangeText, keyword, vbCr & ReplaceInsert)
        End If

        'The disclaimers are after INSERT. Make sure the formatting doesn't change
        With m_wordDoc.Paragraphs(para).Range
            m_wordDoc.Range(.Start, .End).Font.Size = 12
            m_wordDoc.Range(.Start, .End).Font.Italic = False
            m_wordDoc.Range(.Start, .End).Font.Bold = False
        End With
    ElseIf keyword = "Regards" Then
        ReplaceRegards keywordRange, buttonWindow.Book.Worksheets(1)
    End If
    
    'Return whether the keyword was found in the range
    ReplaceKeywords = CBool(keywordRange Is Nothing)
End Function

Private Function FindKeyword(paragraphRange As Word.Range, keyword As String) As Word.Range
    Dim keywordFound As Boolean
    Dim wrd As Integer
    wrd = 1
    Do While wrd < paragraphRange.Words.count And keywordFound = False
        'Get the next word from the paragraph
        Dim currentWordRange As Word.Range
        Set currentWordRange = paragraphRange.Words(wrd)
        Dim currentWord As String
        currentWord = currentWordRange.Text
        
        'See if this word is the keyword
        If currentWord = keyword Then
            Set FindKeyword = currentWordRange
            keywordFound = True
        End If
        
        wrd = wrd + 1
    Loop
End Function

Private Function ReplaceDate() As String
    'Get the current date, the next day's date if it's after 3:00, or the next Monday's date if it's late on Friday
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

Private Function ReplaceDear() As String
    'Get client(s) first name(s)
    Dim fNames As String
    If m_household.SpouseFirst = vbNullString Then
        fNames = m_household.ClientFirst
    Else
        fNames = m_household.ClientFirst & " and " & m_household.SpouseFirst
    End If
    ReplaceDear = "Dear " & fNames
End Function

Private Function ReplaceTarget() As String
    'Add the amount of equities bought or sold
    Dim buysell As Single
    buysell = AmountBoughtSold
    
    'If the buysell is negative, then we're selling equities
    Dim bstStr As String
    If buysell < 0 Then
        'Selling equities
        bstStr = "At this time, we will sell " & Format(-1 * buysell, "$#,##0") _
            & " of stock funds in order to move towards your " & m_household.eqTarget & " equity target"
    ElseIf buysell > 0 Then
        'Buying equities
        bstStr = "At this time, we will buy " & Format(buysell, "$#,##0") _
            & " of stock funds in order to move towards your " & m_household.eqTarget & " equity target"
    Else
        'At target
        bstStr = "At this time you are at your " & m_household.eqTarget & " equity target." _
            & "  We will move money around the investment style grid in order to buy" _
            & " low and sell high among different asset classes"
    End If
    
    ReplaceTarget = bstStr
End Function

Private Function AmountBoughtSold() As Double
    'Get the total bought or sold of FI or MMM
    Dim total As Double
    
    'Loop through each account in the list
    Dim acct As Variant
    For Each acct In m_household.Accounts
        'Look for these subclasses
        Dim fixedArray() As String
        fixedArray = Split("FI;MMM", ";")
        
        'Loop through each transaction in the account
        Dim trans As Variant
        For Each trans In acct.TradeList
            'Loop through each subclass that we're looking for
            Dim fixedSubclass As Variant
            For Each fixedSubclass In fixedArray
                If trans.Subclass = fixedSubclass Then
                    'The transaction is fixed income or money market, add the trade value to the total
                    total = total + trans.Amount
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

Private Function ReplaceRMD(isRMDChecked As Boolean) As String
    If isRMDChecked Then
        'There's at least one RMD. Put in the proper wording here
        'ReplaceRMD =
    Else
        'There's no RMD. Take out the placeholder
        ReplaceRMD = ""
    End If
End Function

Private Function ReplaceInsert() As String
    'This function is trivial, but leaving it in in case something needs to be changed
    ReplaceInsert = m_insert
End Function

Private Sub ReplaceRegards(regardsRange As Word.Range, buttonSheet As Worksheet)
    'Find who's signing the letter
    m_advisors = GetAdvisors(buttonSheet)
    
    'Fill array with advisor names
    Dim advStr As String
    Dim advArray() As String
    advStr = "Dan#Budinger#Ryan#Wempe#Rachel#Brown"
    advArray = Split(advStr, "#")

    'Set a counter for the number of advisor names added
    Dim advNames As Integer
    advNames = 0
    
    'Get the range running from "regards" until the end of the letter
    Dim endRange As Word.Range
    Set endRange = m_wordDoc.Range(regardsRange.Start, m_wordDoc.Range.End)

    'Iterate through the words after regards until an advisor name is found
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
                            endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), m_advisors(0, 0))
                        Else
                            'One advisor has been added, so add the second
                            endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), m_advisors(1, 0))
                        End If
                        advNames = advNames + 1
                    Else
                        'The word is a last name
                        If advNames = 1 Then
                            'No advisors have been added, so add the first
                            endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), m_advisors(0, 1))
                        Else
                            'One advisor has been added, so add the second
                            endRange.Words(wrd).Text = Replace(endRange.Words(wrd).Text, advArray(advIndex), m_advisors(1, 1))
                        End If
                        advNames = advNames + 1
                    End If
                End If
            Next advIndex
        End If
        wrd = wrd + 1
    Loop
End Sub

Private Function GetAdvisors(buttonSheet As Worksheet) As String()
    'Set up a temporary array to return, where the first dimension is the advisor, and the second is first/last name
    Dim advArray(1, 1) As String
    
    'Find which advisor checkboxes are checked
    If buttonSheet.Shapes("cbxDan").OLEFormat.Object.Object.value = True Then
        'If Dan is checked, he will always be the first advisor
        advArray(0, 0) = "Dan"
    End If
    If buttonSheet.Shapes("cbxRyan").OLEFormat.Object.Object.value = True Then
        'Ryan's the first advisor if Dan isn't; otherwise he's the second advisor
        If advArray(0, 0) = vbNullString Then
            advArray(0, 0) = "Ryan"
        Else
            advArray(1, 0) = "Ryan"
        End If
    End If
    If buttonSheet.Shapes("cbxRachel").OLEFormat.Object.Object.value = True Then
        'If Rachel is checked, she will always be the second advisor
        advArray(1, 0) = "Rachel"
    End If
    
    'Get the advisors' last names from their first names
    advArray(0, 1) = GetLast(advArray(0, 0))
    advArray(1, 1) = GetLast(advArray(1, 0))
    
    'Check for errors
    If advArray(0, 0) = vbNullString Then
        'None of the advisors were checked
        ErrorHandling.ErrorAndStop "Please select at least one advisor to sign the letter", "Someone's gotta sign this thing"
    End If
        
    If advArray(0, 1) = vbNullString Or (advArray(1, 0) <> vbNullString And advArray(1, 1) = vbNullString) Then
        'At least one of the last names are missing
        ErrorHandling.ErrorAndContinue "Something went wrong and at least one of the advisors' last names are missing. It will need to be added to the letter manually.", "This was Steve's fault"
    End If
    
    'Return the advisor array
    GetAdvisors = advArray
End Function

Private Function GetLast(advFirst) As String
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

Public Sub PrintLetter()
    'Set the report builder's window
    Dim buttonWindow As New clsWindow
    Set buttonWindow = ClassBuilder.NewWindow("Report Builder")
    
    'Make Word visible and bring it into focus (activate doesn't always do this; not sure why)
    m_wordApp.Visible = True
    m_wordApp.Activate
    
    'See if the RMD box is checked on the report builder
    Dim isRMDChecked As Boolean
    isRMDChecked = buttonWindow.Book.Worksheets(1).Shapes("cbxRMD").OLEFormat.Object.Object.value
    If Not isRMDChecked Then
        'No RMD, print the letter
        LogData.TimeEnd
        m_wordApp.PrintPreview = True
        Do While m_wordApp.PrintPreview = True
            DoEvents
        Loop
        LogData.TimeStart
    Else
        'There's an RMD, don't print the letter so the wording can be put in manually
    End If
End Sub

Public Sub SaveLetter()
    'Default file name is [Month] [Year].docx
    Dim saveFile As String
    saveFile = MonthName(Month(Date)) & " " & Year(Date)
    
    'Open save dialog with the default file name of the current month and current year
    LogData.TimeEnd
    Dim sfdSaveLetter As FileDialog
    Set sfdSaveLetter = m_wordApp.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveLetter.InitialFileName = household.ServerFolder & saveFile
    sfdSaveLetter.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveLetter.SelectedItems.count > 0 Then
        savePath = sfdSaveLetter.SelectedItems.Item(1)
        m_wordDoc.SaveAs2 fileName:=savePath, ReadOnlyRecommended:=False
    End If
    LogData.TimeStart
End Sub

