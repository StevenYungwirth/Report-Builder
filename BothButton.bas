Attribute VB_Name = "BothButton"
Option Explicit
Public ReportBuilderBook As Workbook
Public ReportBuilderSheet As Worksheet
Public ExportBook As Workbook
Public ExportSheet As Worksheet
Public HouseholdName As String
Public saveDir As String
Public EqTarget As String
Public AccountList As Collection
Sub BuildBoth()
    LogData.TimeStart
    UpdateScreen "Off"
    On Error GoTo MacroBroke
    SetGlobals
    
    'Log macro use
    LogData.WriteLog "Both Button", HouseholdName, True
    
    'Build both letter and report and print off each
    ReportsButton.GenerateReports True
    LetterButton.BuildLetter True
    
    ResetGlobals
    UpdateScreen "On"
    
    'Log the time
    LogData.TimeEnd
    LogData.LogTime
    Exit Sub
MacroBroke:
    ErrorHandling.ErrorAndStop
End Sub
Sub UpdateScreen(OnOrOff As String)
    Dim reset As Long
    If OnOrOff = "Off" Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
    ElseIf OnOrOff = "On" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        reset = ActiveSheet.UsedRange.Rows.count
    End If
End Sub
Sub SetGlobals()
    'Set the report builder's workbook and worksheet
    Set ReportBuilderBook = Workbooks(BothButton.BookCheck("Report Builder"))
    Set ReportBuilderSheet = ReportBuilderBook.Worksheets(1)
    
    'Set the worksheet with the trades
    Set ExportBook = Workbooks(BookCheck("TradeRecommendationsExport"))
    Set ExportSheet = ExportBook.Worksheets(1)
    ExportBook.Activate
    ExportSheet.Activate
    
    'Check to make sure the household name is there
    Dim household As Range
    Set household = FindHeader("CRHouseholdDescription")
    
    'Get the household name
    HouseholdName = household.Offset(1, 0).Value2
    
    EqTarget = FindTarget
    saveDir = GetDir
    
    'Get the column headers
    Dim headerStrings() As String
    headerStrings = GetHeaders
    
    'Get each trade row
    Dim tradeRows() As Variant
    tradeRows = GetTrades
    
    'Get List of accounts and their trades
    GetAccounts headerStrings, tradeRows
End Sub
Function GetHeaders() As Variant
    'Get the bounds of the worksheet
    Dim lastCol As Integer
    lastCol = BothButton.ExportSheet.Cells(1, BothButton.ExportSheet.Columns.count).End(xlToLeft).Column
    
    'Set the array to be the values of the first row
    Dim headers() As String
    ReDim headers(0 To lastCol - 1) As String
    Dim col As Integer
    For col = 1 To lastCol
        headers(col - 1) = Trim(BothButton.ExportSheet.Cells(1, col).Value2)
    Next col
    
    'Check to make sure these column headers can be found. If they're not available, an error will be thrown
    GetIndexOf headers, "AccountNumber"
    GetIndexOf headers, "CRAccountMasterDescription"
    GetIndexOf headers, "Custodian"
    GetIndexOf headers, "Symbol"
    GetIndexOf headers, "OriginalTradeDate"
    GetIndexOf headers, "CostBasis"
    GetIndexOf headers, "Trade"
    GetIndexOf headers, "AccountType"
    GetIndexOf headers, "Action"
    GetIndexOf headers, "Description"
    GetIndexOf headers, "PCNTSOLD"
    
    GetHeaders = headers
End Function

Function GetTrades() As Variant
    'Get the bounds of the worksheet
    Dim lastCol As Integer
    lastCol = BothButton.ExportSheet.Cells(1, BothButton.ExportSheet.Columns.count).End(xlToLeft).Column
    Dim lastRow As Integer
    lastRow = BothButton.ExportSheet.Cells(BothButton.ExportSheet.Rows.count, 1).End(xlUp).Row
    
    'Get each trade row
    Dim tempArr() As Variant
    ReDim tempArr(0 To lastRow - 2, 0 To lastCol - 1) As Variant
    Dim rw As Integer
    Dim col As Integer
    For rw = 2 To lastRow
        For col = 1 To lastCol
            tempArr(rw - 2, col - 1) = Trim(BothButton.ExportSheet.Cells(rw, col).Value2)
        Next col
    Next rw
    GetTrades = tempArr
End Function

Sub GetAccounts(headerStrings As Variant, tradeRows As Variant)
    'Define a new collection
    Set AccountList = New Collection
    
    'Loop through each trade row element, get each account and their respective trades
    Dim ele As Integer
    For ele = 0 To UBound(tradeRows, 1)
        'Set the trade row as a new trade row object
        Dim tempRow As clsTradeRow
        Set tempRow = New clsTradeRow
        tempRow.Symbol = UCase(tradeRows(ele, GetIndexOf(headerStrings, "Symbol")))
        tempRow.Description = tradeRows(ele, GetIndexOf(headerStrings, "Description"))
        tempRow.subclass = tradeRows(ele, GetIndexOf(headerStrings, "SubClass"))
        tempRow.Action = tradeRows(ele, GetIndexOf(headerStrings, "Action"))
        
        'If the trade/percent cell is empty, return 0 (trade should never be empty, but this is just in case)
        If tradeRows(ele, GetIndexOf(headerStrings, "Trade")) = vbNullString Then
            tempRow.Trade = 0
        Else
            tempRow.Trade = tradeRows(ele, GetIndexOf(headerStrings, "Trade"))
        End If
        If tradeRows(ele, GetIndexOf(headerStrings, "PCNTSOLD")) = vbNullString Then
            tempRow.Percent = 0
        Else
            tempRow.Percent = tradeRows(ele, GetIndexOf(headerStrings, "PCNTSOLD"))
        End If
        
        'Loop through each account in the account list to see if it's already there
        Dim isAccountInList As Boolean
        isAccountInList = False
        Dim acct As Variant
        For Each acct In AccountList
            If tradeRows(ele, GetIndexOf(headerStrings, "AccountNumber")) = acct.Number Then
                'Account is already in the list, add the trade row to the account
                acct.TradeList.Add tempRow
                isAccountInList = True
            End If
        Next acct
        
        If Not isAccountInList Then
            'The account isn't in the list, create a new account
            Dim tempAccount As clsAccount
            Set tempAccount = New clsAccount
            tempAccount.household = tradeRows(ele, GetIndexOf(headerStrings, "CRHouseholdDescription"))
            tempAccount.EqTarget = tradeRows(ele, GetIndexOf(headerStrings, "AssetAllocationModel"))
            tempAccount.Number = tradeRows(ele, GetIndexOf(headerStrings, "AccountNumber"))
            tempAccount.Name = tradeRows(ele, GetIndexOf(headerStrings, "CRAccountMasterDescription"))
            tempAccount.AcctType = tradeRows(ele, GetIndexOf(headerStrings, "AccountType"))
            tempAccount.Custodian = tradeRows(ele, GetIndexOf(headerStrings, "Custodian"))
            
            'Add the trade row to a new account
            tempAccount.TradeList.Add tempRow
            
            'Add the new account to the list
            AccountList.Add tempAccount
        End If
    Next ele
End Sub
Sub ResetGlobals()
    Set ReportBuilderSheet = Nothing
    Set ReportBuilderBook = Nothing
    Set ExportSheet = Nothing
    Set ExportBook = Nothing
    LogData.LogStarted = False
End Sub
Function BookCheck(fileName) As String
    Dim numberOfWindows As Integer
    Dim window As Integer
    Dim windowName As String
    
    numberOfWindows = Windows.count
    For window = 1 To numberOfWindows
        windowName = Windows(window).Caption
        If InStr(UCase(windowName), UCase(fileName)) > 0 Then
            BookCheck = Windows(window).Caption
        End If
    Next window
    
    If BookCheck = vbNullString Then
        ErrorHandling.ErrorAndStop "Recommendations sheet not found"
    End If
End Function
Function FindHeader(Target As String) As Range
    Dim rng As Range
    Set rng = ExportSheet.UsedRange.Find(Target, After:=ExportSheet.Range("A1"), LookAt:=xlWhole)
    If rng Is Nothing Then
        Set rng = ExportSheet.UsedRange.Find(" " & Target, After:=ExportSheet.Range("A1"), LookAt:=xlWhole)
    End If
    
    If rng Is Nothing Then
        ErrorHandling.ErrorAndStop Target & " not found, macro has halted"
    Else
        Set FindHeader = rng
    End If
End Function
Function GetDir()
    'Set the folder name
    Dim folderName As String
    folderName = ClientFolder
    
    'If the client's folder is found, set the default directory to be that
    GetDir = "Z:\" & folderName & "\"
    If Dir(GetDir, vbDirectory) <> "" Then
        'The folder exists, see if it has a letters folder
        If Dir(GetDir & "Letters" & "\", vbDirectory) <> "" Then
            'Full path is available, have the default save location be here
            GetDir = GetDir & "Letters" & "\"
        ElseIf Dir(GetDir & "Letter" & "\", vbDirectory) <> "" Then
            'Full path is available, have the default save location be here
            GetDir = GetDir & "Letter" & "\"
        Else
            'Have the default save location be the client's folder
        End If
    Else
        'The client's folder can't be found, open the dialog at the Z drive
        GetDir = "Z:\"
    End If
End Function
Function ClientFolder() As String
    Dim folderName As String
    folderName = HouseholdName
    If FolderExists(folderName) Then
        'Good; folder name is household name
    Else
        If InStr(HouseholdName, ",") > 0 Then
            'Household name has a comma, find where it is
            Dim commaPos As Integer
            commaPos = InStr(HouseholdName, ",")
            
            'Declare name arrays
            Dim clientName(2) As String
            Dim spouseName(2) As String
            
            'Take last name as everything before the comma
            clientName(1) = Trim(Left(HouseholdName, commaPos - 1))
            clientName(1) = Replace(clientName(1), " ", "")
            clientName(1) = Replace(clientName(1), "-", "")
            
            'Get first name
            If InStr(HouseholdName, "&") > 0 Then
                'Household name has spouse's name
                Dim ampPos As Integer
                ampPos = InStr(HouseholdName, "&")

                'Everything between comma and ampersand is the first name
                clientName(0) = Trim(Mid(HouseholdName, commaPos + 1, ampPos - commaPos - 1))
                
                'Everything after the ampersand is the spouse's first name
                spouseName(0) = Trim(Right(HouseholdName, Len(HouseholdName) - ampPos - 1))
                
                If InStr(spouseName(0), ",") > 0 Then
                    'Spouse has a different last name
                    Dim spouseCommaPos As Integer
                    spouseCommaPos = InStr(spouseName(0), ",")
                    spouseName(1) = Trim(Left(spouseName(0), spouseCommaPos - 1))
                    spouseName(1) = Replace(spouseName(1), " ", "")
                    spouseName(1) = Replace(spouseName(1), "-", "")
                    spouseName(0) = Trim(Right(spouseName(0), Len(spouseName(0)) - spouseCommaPos))
                Else
                    'Spouse has the same last name
                    spouseName(1) = clientName(1)
                    spouseName(1) = Replace(spouseName(1), " ", "")
                    spouseName(1) = Replace(spouseName(1), "-", "")
                End If
            Else
                'Household name is just one person, take everything after the comma as the first name
                clientName(0) = Trim(Right(HouseholdName, Len(HouseholdName) - commaPos))
                spouseName(0) = ""
                spouseName(1) = ""
            End If
            
            If InStr(clientName(0), " ") > 0 Then
                'Client's name has a suffix
                clientName(2) = Trim(Right(clientName(0), Len(clientName(0)) - InStr(clientName(0), " ")))
                clientName(0) = Trim(Left(clientName(0), InStr(clientName(0), " ")))
            End If
            If InStr(spouseName(0), " ") > 0 Then
                'Spouse's name has a suffix
                spouseName(2) = Trim(Right(spouseName(0), Len(spouseName(0)) - InStr(spouseName(0), " ")))
                spouseName(0) = Trim(Left(spouseName(0), InStr(spouseName(0), " ")))
            End If
                
            'Get the name of the client's folder. Returns folderName="" if not found
            folderName = FindFolder(clientName, spouseName)
        Else
            'Household name has no comma, and the household name isn't the folder
        End If
    End If
    
    ClientFolder = folderName
End Function
Function FolderExists(dirStr As String) As Boolean
    'Return whether or not the folder name is found in the Z drive
    Dim saveDir As String
    saveDir = "Z:\" & dirStr & "\"
    If Dir(saveDir, vbDirectory) = "" Or dirStr = "" Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function
Function FindFolder(clientName() As String, spouseName() As String) As String
    Dim folderName As String
    
    'Both names are [first name][last name][suffix]
    Dim cName As String
    Dim sName As String
    cName = clientName(1) & clientName(0) & clientName(2)
    sName = spouseName(1) & spouseName(0) & spouseName(2)
    
    'Find the first letter of the client's name
    Dim firstChr As String
    firstChr = UCase(Left(cName, 1))
    
    'Get the first folder that starts with the first letter
    Dim zFolders() As String
    ReDim zFolders(0 To 0)
    zFolders(0) = Dir("Z:\" & firstChr & "*", vbDirectory)
    Dim count As Integer
    
    'Get the next folder that starts with the first letter
    Dim nextDir As String
    nextDir = Dir()
    
    'Get all folder names that start with the first letter of the client's name
    count = 0
    Do While nextDir <> ""
        count = count + 1
        ReDim Preserve zFolders(0 To count)
        zFolders(count) = nextDir
        nextDir = Dir()
    Loop
    
    'Narrow down the array by starting with the first two letters of the client's last name
    'Reduce the options until the folder is found or the end of the name is reached
    Dim chr As Integer
    chr = 2
    
    Dim exitLoop As Boolean
    exitLoop = False
    Do While Not exitLoop And chr <= Len(cName)
        zFolders = ReduceArr(cName, zFolders, chr)
        If UBound(zFolders) = 0 Then
            'The array only has one entry left, no need to go further
            folderName = zFolders(0)
            exitLoop = True
        End If
        chr = chr + 1
    Loop
    
    'Array could have multiple values in it after this
    If UBound(zFolders) > 0 Then
        If spouseName(0) <> "" Then
            'Folder isn't under client's name, try the spouse's name
            Dim emptyArr(2) As String
            folderName = FindFolder(spouseName, emptyArr)
        Else
            'Spouse's name was tried already
        End If
    End If
    
    If Not FolderExists(folderName) Then
        folderName = NameCombinations(clientName(1), clientName(0), spouseName(0), spouseName(1))
    End If
    FindFolder = folderName
End Function
Function OnlyLetters(str As String) As String
    Dim i As Integer
    Dim tempStr As String
    tempStr = ""
    For i = 1 To Len(str)
        If IsLetter(Mid(str, i, 1)) Then
            tempStr = tempStr & Mid(str, i, 1)
        End If
    Next i
    OnlyLetters = tempStr
End Function
Function IsLetter(r As String) As Boolean
    Dim x As String
    x = UCase(r)
    IsLetter = Asc(x) > 64 And Asc(x) < 91
End Function
Function ReduceArr(Name As String, arr() As String, chr As Integer) As String()
    Dim xChr As String
    xChr = UCase(Left(Name, chr))
    
    Dim outArr() As String
    Dim xCount As Integer
    Dim ele As Variant
    Dim eleName As String
    ReDim outArr(0) As String
    For Each ele In arr
        If Left(UCase(ele), chr) = xChr Then
            ReDim Preserve outArr(0 To xCount) As String
            outArr(xCount) = ele
            xCount = xCount + 1
        End If
    Next ele

    ReduceArr = outArr
End Function
Function NameCombinations(lName As String, fName As String, Optional spouseFirst As String, Optional spouseLast As String) As String
    'Declare arrays for every possible character or word in every position
    Dim posZero() As String
    Dim posOne() As String
    Dim posTwo() As String
    Dim posThree() As String
    Dim posFour() As String
    Dim posFive() As String
    Dim posSix() As String
    Dim posSeven() As String
    
    'Not as many options for the folder name if there's only one person
    If spouseFirst = "" Then
        ReDim posFour(0) As String
        ReDim posFive(0) As String
        ReDim posSix(0) As String
        ReDim posSeven(0) As String
        'last
        posZero = Split(lName, ";")
        'blank, space, comma, first
        posOne = Split("; ;,;" & fName, ";")
        'blank, space, first
        posTwo = Split("; ;" & fName, ";")
        'blank, first
        posThree = Split(";" & fName, ";")
        'blank
        posFour(0) = ""
        'blank
        posFive(0) = ""
        'blank
        posSix(0) = ""
        'blank
        posSeven(0) = ""
    Else
        'last, spouse last
        posZero = Split(lName & ";" & spouseLast, ";")
        'blank, space, comma, first, spouse first, spouse last
        posOne = Split("; ;,;" & fName & ";" & spouseFirst & ";" & spouseLast, ";")
        'blank, space, &, and, first, spouse first
        posTwo = Split("; ;&;and;" & fName & ";" & spouseFirst, ";")
        'blank, space, &, and, first, spouse first
        posThree = Split("; ;&;and;" & fName & ";" & spouseFirst, ";")
        'blank, space, &, and, first, spouse first
        posFour = Split("; ;&;and;" & fName & ";" & spouseFirst, ";")
        'blank, space, &, and, first, spouse first
        posFive = Split("; ;&;and;" & fName & ";" & spouseFirst, ";")
        'blank, space, first, spouse first
        posSix = Split("; ;" & fName & ";" & spouseFirst, ";")
        'blank, first, spouse first
        posSeven = Split(";" & fName & ";" & spouseFirst, ";")
    End If
    
    'Run through every combination of possible folder names, but leave the loops once the folder name is found
    Dim folderName As String
    Dim found As Boolean
    Dim zeroth, first, second, third, fourth, fifth, sixth, seventh As Integer
    found = False
    
    seventh = 0
    Do While Not found And seventh <= UBound(posSeven)
        sixth = 0
        Do While Not found And sixth <= UBound(posSix)
            fifth = 0
            Do While Not found And fifth <= UBound(posFive)
                fourth = 0
                Do While Not found And fourth <= UBound(posFour)
                    third = 0
                    Do While Not found And third <= UBound(posThree)
                        second = 0
                        Do While Not found And second <= UBound(posTwo)
                            first = 0
                            Do While Not found And first <= UBound(posOne)
                                zeroth = 0
                                Do While Not found And zeroth <= UBound(posZero)
                                    folderName = posZero(zeroth) & posOne(first) & posTwo(second) & posThree(third) & posFour(fourth) & posFive(fifth) & posSix(sixth) & posSeven(seventh)
                                    found = FolderExists(folderName)
                                    zeroth = zeroth + 1
                                Loop
                                first = first + 1
                            Loop
                            second = second + 1
                        Loop
                        third = third + 1
                    Loop
                    fourth = fourth + 1
                Loop
                fifth = fifth + 1
            Loop
            sixth = sixth + 1
        Loop
        seventh = seventh + 1
    Loop
    
    If Not found Then
        NameCombinations = ""
    Else
        NameCombinations = folderName
    End If
End Function
Function FindTarget() As String
    Dim equityTargetLoc As Range
    Dim equityStr As String
    Set equityTargetLoc = FindHeader("AssetAllocationModel")
    equityStr = equityTargetLoc.Offset(1, 0).Value
    
    'If there's a % in the string, remove it and everything to the right of it
    If InStr(equityStr, "%") <> 0 Then
        equityStr = Left(equityStr, InStr(equityStr, "%") - 1)
    End If
    
    'Remove characters until eqTarget is either blank or numeric
    Do While Not IsNumeric(equityStr) And equityStr <> ""
        If InStr(equityStr, " ") > 1 Then
            'There is a space
            If Not IsNumeric(Mid(equityStr, InStr(equityStr, " ") - 1, 1)) Then
                'The character before the space is not numeric, so remove it
                equityStr = Right(equityStr, Len(equityStr) - 1)
            Else
                'The character before the space is numeric
                If InStr(equityStr, "-") > 0 Or InStr(equityStr, "/") > 0 Then
                    If InStr(equityStr, "-") > InStr(equityStr, " ") Or InStr(equityStr, "/") > InStr(equityStr, " ") Then
                        'There's a special character after the space. Remove characters at the end until there's no space
                        Do While InStr(equityStr, " ") > 0
                            equityStr = Left(equityStr, Len(equityStr) - 1)
                        Loop
                    Else
                        'There's a special character before the space. Remove characters at the start until there's no space
                        Do While InStr(equityStr, " ") > 0
                            equityStr = Right(equityStr, Len(equityStr) - 1)
                        Loop
                    End If
                Else
                    'Default: The character before the space is numeric, but there's no special character. Remove characters at the end
                    equityStr = Left(equityStr, Len(equityStr) - 1)
                End If
            End If
        ElseIf InStr(equityStr, " ") = 1 Then
            'The space is the first character, so remove it
            equityStr = Right(equityStr, Len(equityStr) - 1)
        Else
            'No space, e.g. fpis80/80fpis
            If IsNumeric(Left(equityStr, 1)) Then
                'The first character is numeric, remove characters at the end
                equityStr = Left(equityStr, Len(equityStr) - 1)
            Else
                'The first character is not numeric, remove characters from the start
                equityStr = Right(equityStr, Len(equityStr) - 1)
            End If
        End If
    Loop
    
    If equityStr = "" Then
        'equityStr (presumably) contained only text
        FindTarget = "N/A"
    Else
        FindTarget = Trim(equityStr) & "%"
    End If
End Function
