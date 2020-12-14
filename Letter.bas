Attribute VB_Name = "Letter"
Option Explicit
Private buttonBook As Workbook
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
    'Add check: if there's an RMD and it's the months we add RMD stuff, have a message box saying to add something to the letter; don't print the letter
    'Set global variables
    Both.UpdateScreen "Off"
    'On Error GoTo MacroBroke
    If Not bothClicked Then
        Both.SetGlobals
    End If
    SetLetterGlobals
    
    'Open the letter
    OpenLetter
    
    'Add client names and advisor names to letter
    ProcessLetter
    
    'Print the letter
    PrintLetter
    
    'Save the letter
    SaveLetter
    
    ResetLetterGlobals
    If Not bothClicked Then
        Both.ResetGlobals
    End If
    Both.UpdateScreen "On"
    Exit Sub
MacroBroke:
    If Not WordDoc Is Nothing Then
        WordDoc.Close SaveChanges:=False
    End If
    If Not WordApp Is Nothing Then
        WordApp.Quit
    End If
    Both.UpdateScreen "On"
    MsgBox "Fatal error, macro has halted"
End Sub
Sub SetLetterGlobals()
    'Set the button's workbook and worksheet
    Set buttonBook = Workbooks(Both.BookCheck("Report Builder"))
    Set buttonSheet = buttonBook.Worksheets(1)
    
    'Find who's signing the letter
    GetAdvisors
End Sub
Sub ResetLetterGlobals()
    Set buttonSheet = Nothing
    Set buttonBook = Nothing
    Set WordDoc = Nothing
    Set WordApp = Nothing
End Sub
Sub GetAdvisors()
adv1First = ""
adv1Last = ""
adv2First = ""
adv2Last = ""

'Find which advisor checkboxes are checked
If buttonSheet.Shapes("cbxDan").OLEFormat.Object.Object.Value = True Then
    adv1First = "Dan"
    adv1Last = "Budinger"
End If
If buttonSheet.Shapes("cbxRyan").OLEFormat.Object.Object.Value = True Then
    If adv1First = "" Then
        adv1First = "Ryan"
        adv1Last = "Wempe"
    Else
        adv2First = "Ryan"
        adv2Last = "Wempe"
    End If
End If
If buttonSheet.Shapes("cbxRachel").OLEFormat.Object.Object.Value = True Then
    adv2First = "Rachel"
    adv2Last = "Brown"
End If

If adv1First = "" Then
    MsgBox "Please select at least one advisor to sign the letter"
    Both.UpdateScreen "On"
    End
End If
    
If adv1Last = "" Or (adv2First <> "" And adv2Last = "") Then
    'At least one of the last names are missing
    MsgBox "At least one of the advisors' last names are missing. It will need to be added to the letter manually."
End If
End Sub
Function GetLast(advFirst) As String
    If advFirst = "Dan" Then
        GetLast = "Budinger"
    ElseIf advFirst = "Ryan" Then
        GetLast = "Wempe"
    ElseIf advFirst = "Rachel" Then
        GetLast = "Brown"
    End If
End Function
Sub OpenLetter()
    Const defDir As String = "Z:\"
    Const defExtension As String = ".docx"
    
    'Build letterPath using a directory, folder, file, and extension
    'Use provided file if available. If not, use the default file. If that's not available, open a file dialog
    Dim letterPath As String
    Dim folder As String
    Dim file As String
    Dim extension As String
    
    'Use provided file if available
    'Get the file
    If Not buttonSheet.UsedRange.Find("File", After:=buttonSheet.Range("A1"), LookAt:=xlWhole) Is Nothing Then
        file = buttonSheet.UsedRange.Find("File", After:=buttonSheet.Range("A1"), LookAt:=xlWhole).Offset(0, 1).Value
        If file = "" Then
            file = defFile
        End If
    Else
        file = defFile
    End If
    
    'Get the folder
    If Not buttonSheet.UsedRange.Find("Folder", After:=buttonSheet.Range("A1"), LookAt:=xlWhole) Is Nothing Then
        folder = buttonSheet.UsedRange.Find("Folder", After:=buttonSheet.Range("A1"), LookAt:=xlWhole).Offset(0, 1).Value
        If folder = "" Then
            folder = defFolder
        Else
            'Ensure the folder ends with a \
            If Right(folder, 1) <> "\" Then
                folder = folder & "\"
            End If
        End If
    Else
        folder = defFolder
    End If
    
    'Get the extension
    If Not buttonSheet.UsedRange.Find("Extension", After:=buttonSheet.Range("A1"), LookAt:=xlWhole) Is Nothing Then
        extension = buttonSheet.UsedRange.Find("Extension", After:=buttonSheet.Range("A1"), LookAt:=xlWhole).Offset(0, 1).Value
        If extension = "" Then
            extension = defExtension
        End If
    Else
        extension = defExtension
    End If
    
    letterPath = defDir & folder & file & extension
    
    'If the provided file doesn't exist, use the default file
    If Dir(letterPath) = "" Then
        letterPath = defDir & defFolder & defFile & defExtension
    End If
    
    'If the default location doesn't exist, select a file to open
    If Dir(letterPath) = "" Then
        letterPath = FindPath
    End If
 
    'Open the file
    Set WordApp = New Word.Application
    'Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False
    Set WordDoc = WordApp.Documents.Open(letterPath, ReadOnly:=True)
    WordDoc.Activate
    WordDoc.ActiveWindow.View.ReadingLayout = False
End Sub
Function defFile()
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
    defFile = quarter & " Quarter " & Year(Date) & " - " & quarterMonths
End Function
Function defFolder()
    defFolder = "FPIS - Operations\Client Communications\FPIS - Inv Rec Letters\"
End Function
Function FindPath() As String
    'Get a file
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Word Files", "*.docx; *.doc", 1
        .InitialFileName = "Z:\FPIS - Operations\Client Communications\FPIS - Inv Rec Letters\"
        .Show
        
        Dim fileSelected As Variant
        For Each fileSelected In .SelectedItems
            FindPath = .SelectedItems.Item(1)
        Next fileSelected
    End With
    
    'Quit the procedure if the user didn't select the type of file we need.
    If InStr(FindPath, ".doc") = 0 Then
        Exit Function
    End If
End Function
Sub ProcessLetter() 'Iterate through letter and add the client(s) name(s) and advisor names
    'Get client(s) first name(s)
    Dim hhname As String
    Dim fNames As String
    hhname = household.Offset(1, 0).Value
    
    If InStr(hhname, ",") > 0 Then
        'Household name has a comma, find where it is
        Dim commaPos As Integer
        commaPos = InStr(hhname, ",")
        
        'Get first name
        Dim fName As String
        Dim spouseFirst As String
        If InStr(hhname, "&") > 0 Then
            'Household name has spouse's name
            Dim ampPos As Integer
            ampPos = InStr(hhname, "&")

            'Everything between comma and ampersand is the first name
            fName = Trim(Mid(hhname, commaPos + 1, ampPos - commaPos - 1))
            fName = Replace(fName, "Jr", "")
            fName = Trim(Replace(fName, "Sr", ""))
            
            'Everything after the ampersand is the spouse's first name
            spouseFirst = Trim(Right(hhname, Len(hhname) - ampPos - 1))
            
            If InStr(spouseFirst, ",") > 0 Then
                'Spouse has a different last name, take everything after the second comma
                Dim spouseCommaPos As Integer
                spouseCommaPos = InStr(spouseFirst, ",")
                spouseFirst = Trim(Right(spouseFirst, Len(spouseFirst) - spouseCommaPos))
            End If
            fNames = fName & " and " & spouseFirst
        Else
            'Household name is just one person, take everything after the comma as the first name
            fName = Trim(Right(hhname, Len(hhname) - commaPos))
            fName = Replace(fName, "Jr", "")
            fName = Trim(Replace(fName, "Sr", ""))
            fNames = fName
        End If
    End If
            
    'Iterate through letter by word
    Dim wrd As Integer
    Dim currentWord As String
    Dim fndDate As Boolean
    Dim fndDear As Boolean
    Dim fndTarget As Boolean
    Dim fndRMD As Boolean
    Dim fndInsert As Boolean
    Dim fndClosing As Boolean
    Dim fndRegards As Boolean
    fndDate = False
    fndDear = False
    fndTarget = False
    fndInsert = False
    fndClosing = False
    fndRegards = False
    
    wrd = 1
    With WordDoc
        Do While Not (fndDear And fndTarget And fndRegards) And wrd < .Words.count
            currentWord = Trim(.Words(wrd).Text)
            If Not fndDate And currentWord = "DATE" Then
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
                
                .Words(wrd) = Replace(.Words(wrd), "DATE", currentDate)
                fndDate = True
            ElseIf Not fndDear And currentWord = "Dear" Then
                'Add the clients' first names after "Dear"
                If currentWord = "Dear " Then
                    .Words(wrd) = Replace(.Words(wrd), "Dear ", "Dear " & fNames)
                Else
                    .Words(wrd) = Replace(.Words(wrd), "Dear", "Dear " & fNames)
                End If
                fndDear = True
            ElseIf Not fndTarget And currentWord = "BUYSELLTARGET" Then
                'Add the amount of equities bought or sold
                Dim buysell As Single
                buysell = BoughtSold
                
                'If the buysell is negative, then we're selling equities
                Dim bstStr As String
                If buysell < 0 Then
                    'Selling equities
                    bstStr = "At this time, we will sell " & Format(-1 * buysell, "$#,##0") _
                        & " of stock funds in order to move towards your " & eqTarget & " equity target"
                ElseIf buysell > 0 Then
                    'Buying equities
                    bstStr = "At this time, we will buy " & Format(buysell, "$#,##0") _
                        & " of stock funds in order to move towards your " & eqTarget & " equity target"
                Else
                    'At target
                    bstStr = "At this time you are at your " & eqTarget & " equity target." _
                        & "  We will move money around the investment style grid in order to buy" _
                        & " low and sell high among different asset classes"
                End If
                
                If currentWord = "BUYSELLTARGET " Then
                    .Words(wrd) = Replace(.Words(wrd), "BUYSELLTARGET ", bstStr)
                Else
                    .Words(wrd) = Replace(.Words(wrd), "BUYSELLTARGET", bstStr)
                End If
                fndTarget = True
            ElseIf Not fndInsert And currentWord = "INSERT" Then
                fndInsert = True
                
                Dim insertLoc As Integer
                insertLoc = .Range(0, .Words(wrd).Start).Paragraphs.count + 1
                .Words(wrd) = Replace(.Words(wrd), "INSERT", "")
                
                Dim insert As String
                Dim paragraph As Word.paragraph
                If buttonSheet.Shapes("rdoTLH").OLEFormat.Object.Object.Value Then
                    insert = "We are making a couple extra changes in your taxable account to harvest some losses.  " _
                    & "These efforts help us offset gains in the future and provide you with a small tax deduction (over any above gains).  " _
                    & "We have done this in the past and found it to be effective."
                    
                    Set paragraph = .Paragraphs.Add(.Paragraphs(insertLoc).Range)
                    paragraph.Range.Text = vbCr & insert & vbCr
                
                    .Range(.Paragraphs(insertLoc).Range.Start, .Paragraphs(insertLoc + 1).Range.End).Font.Size = 12
                    .Range(.Paragraphs(insertLoc).Range.Start, .Paragraphs(insertLoc + 1).Range.End).Font.Italic = False
                ElseIf buttonSheet.Shapes("rdoWD").OLEFormat.Object.Object.Value Then
                    insert = "Since you are drawing on your investments regularly, we are going to configure our rebalancing efforts to adjust for future withdrawals (over the next 3-6 months).  " _
                    & "As you know, we use our most stable bond funds to accommodate withdrawals during market pull backs.  " _
                    & "If you want to make any changes to your withdrawals in the short-term, please let us know."
                    
                    Set paragraph = .Paragraphs.Add(.Paragraphs(insertLoc).Range)
                    paragraph.Range.Text = vbCr & insert & vbCr
                
                    .Range(.Paragraphs(insertLoc).Range.Start, .Paragraphs(insertLoc + 1).Range.End).Font.Size = 12
                    .Range(.Paragraphs(insertLoc).Range.Start, .Paragraphs(insertLoc + 1).Range.End).Font.Italic = False
                End If
            ElseIf Not fndClosing And currentWord = "CLOSING" Then
                fndClosing = True
                .Words(wrd) = Replace(.Words(wrd), "CLOSING ", "")
'                If buttonSheet.Shapes("cbxRMD").OLEFormat.Object.Object.Value Then
'                    .Words(wrd) = Replace(.Words(wrd), "CLOSING ", "")
'                ElseIf .Words(wrd).Information(wdActiveEndPageNumber) = 2 Then
'                    .Words(wrd).InsertBreak (wdPageBreak)
'                End If
            ElseIf Not fndRegards And currentWord = "Regards" Then
                'Add the advisor names, which appear after "Regards,"
                'Fill array with advisor names
                Dim advStr As String
                Dim advArray() As String
                advStr = "Dan#Budinger#Ryan#Wempe#Rachel#Brown"
                advArray = Split(advStr, "#")
    
                Dim advNames As Integer
                advNames = 0
    
                'Loop through the words after regards until an advisor name is found
                Dim endWrd As Integer
                endWrd = wrd
                Do While endWrd < .Words.count - 1 And advNames < 4
                    currentWord = Trim(.Words(endWrd).Text)
                    
                    'If the word is in the advisor array
                    If InStr(advStr, currentWord) > 0 Then
                        Dim advIndex As Integer
                        For advIndex = 0 To UBound(advArray)
                            If currentWord = advArray(advIndex) Then
                                'The word is in the advisor array
                                If advIndex Mod 2 = 0 Then
                                    'The word is a first name
                                    If advNames = 0 Then
                                        'No advisors have been added, so add the first
                                        .Words(endWrd).Text = Replace(.Words(endWrd).Text, advArray(advIndex), adv1First)
                                    Else
                                        'One advisor has been added, so add the second
                                        .Words(endWrd).Text = Replace(.Words(endWrd).Text, advArray(advIndex), adv2First)
                                    End If
                                    advNames = advNames + 1
                                Else
                                    'The word is a last name
                                    If advNames = 1 Then
                                        'No advisors have been added, so add the first
                                        .Words(endWrd).Text = Replace(.Words(endWrd).Text, advArray(advIndex), adv1Last)
                                    Else
                                        'One advisor has been added, so add the second
                                        .Words(endWrd).Text = Replace(.Words(endWrd).Text, advArray(advIndex), adv2Last)
                                    End If
                                    advNames = advNames + 1
                                End If
                            End If
                        Next advIndex
                    End If
                    endWrd = endWrd + 1
                Loop
                fndRegards = True
            End If
            wrd = wrd + 1
        Loop
    End With
End Sub
Function BoughtSold() As Single
    'In the subclass column of the trade recommendations export, find each instance of FI or MMM
    Dim total As Single
    total = 0
    
    'Get the headers needed
    Dim subclass As Range
    Dim trade As Range
    Set subclass = Both.FindHeader("SubClass")
    Set trade = Both.FindHeader("Trade")
    
    'Look for these subclasses
    Dim fixedArray() As String
    fixedArray = Split("FI;MMM", ";")
    
    'Find the last row to see how many times to iterate
    Dim lastRow As Integer
    lastRow = Source.UsedRange.Rows.count
    
    'Look through each transaction
    Dim trans As Integer
    Dim index As Integer
    For trans = 1 To lastRow
        For index = 0 To UBound(fixedArray)
            'See if the transaction is fixed income or money market
            If subclass.Offset(trans).Value2 = fixedArray(index) Then
                'The transaction is fixed income or money market, add the trade value to the total
                total = total + trade.Offset(trans).Value2
            End If
        Next index
    Next trans
    
    'Since this is the total fixed income bought or sold, take the opposite of it to get the equities
    total = -1 * total
    
    'Round to the nearest 100
    total = Round(total / 100) * 100
    
    'Return the overall amount of equities bought or sold in the entire portfolio
    BoughtSold = total
End Function
Sub PrintLetter()
    WordApp.Visible = True
    WordApp.Activate
    
    If buttonSheet.Shapes("cbxRMD").OLEFormat.Object.Object.Value = False Then
        'No RMD, print the letter
        WordApp.PrintPreview = True
        Do While WordApp.PrintPreview = True
            DoEvents
        Loop
    Else
        'There's an RMD, don't print the letter so the wording can be put in manually
    End If
End Sub
Sub SaveLetter()
    'Fill array with names of months
    Dim months() As String
    months = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
    
    'Default file name is [Month] [Year].docx
    Dim saveFile As String
    saveFile = months(Month(Date) - 1) & " " & Year(Date)
    
    'Open save dialog with the default file name of the current month and current year
    Dim sfdSaveLetter As FileDialog
    Set sfdSaveLetter = WordApp.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveLetter.InitialFileName = saveDir & saveFile
    sfdSaveLetter.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveLetter.SelectedItems.count > 0 Then
        savePath = sfdSaveLetter.SelectedItems.Item(1)
        WordDoc.SaveAs2 fileName:=savePath, ReadOnlyRecommended:=False
    End If
End Sub
