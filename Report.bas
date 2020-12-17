Option Explicit
Private ClientSheet As Worksheet
Private HeaderStrings() As Variant
Private TradeRows() As Variant
Private AccountList As Collection
Private CurrentAcct As clsAccount

Sub TestBuildReport(Optional bothClicked As Boolean)
    'Take exported csv from TRX and build the client's report from it
    
    'Turn off screen updating to improve performance
    Both.UpdateScreen "Off"
    
    'Add vague error tracking
    On Error GoTo MacroBroke
    
    'Set global variables
    If Not bothClicked Then
        'Just the report is being generated, not both report and letter. Set the global variables
        '(TradeRecommendationsExport workbook and worksheet, household name, equity target, save location)
        Both.SetGlobals
    End If
    SetReportGlobals
    
    'Process each account
    Dim acct As Variant
    For Each acct In AccountList
        'Add any cash trades to the money market trades, combine sells from different trade lots, and sort the transactions
        Set CurrentAcct = acct
        ProcessAccount
    Next acct
    
    'Generate the report
    GenerateReport

    'Print the report
    PrintReport
    
    'Save the exported csv as an excel workbook in the client's folder
    SaveReport
    
    'Return everything back to normal
    Set ClientSheet = Nothing
    If Not bothClicked Then
        Both.ResetGlobals
    End If
    Both.UpdateScreen "On"
    Exit Sub
MacroBroke:
    Both.UpdateScreen "On"
    MsgBox "Fatal error, macro has halted"
End Sub

Sub SetReportGlobals()
    'Get the column headers
    GetHeaders
    
    'Get each trade row
    GetTrades
    
    'Get List of accounts and their trades
    GetAccounts
End Sub

Sub GetHeaders()
    'Get the bounds of the worksheet
    Dim lastCol As Integer
    lastCol = Source.Cells(1, Source.Columns.count).End(xlToLeft).Column
    
    'Set the array to be the values of the first row
    Dim tempArr() As Variant
    tempArr = Array(Source.Range(Source.Cells(1, 1), Source.Cells(1, lastCol)).Value)
    HeaderStrings = tempArr(0)
    
    'Check to make sure these column headers can be found. If they're not available, an error will be thrown
    GetIndexOf HeaderStrings, "AccountNumber"
    GetIndexOf HeaderStrings, "CRAccountMasterDescription"
    GetIndexOf HeaderStrings, "Custodian"
    GetIndexOf HeaderStrings, "Symbol"
    GetIndexOf HeaderStrings, "OriginalTradeDate"
    GetIndexOf HeaderStrings, "CostBasis"
    GetIndexOf HeaderStrings, "Trade"
    GetIndexOf HeaderStrings, "AccountType"
    GetIndexOf HeaderStrings, "Action"
    GetIndexOf HeaderStrings, "Description"
    GetIndexOf HeaderStrings, "PCNTSOLD"
End Sub

Sub GetTrades()
    'Get the bounds of the worksheet
    Dim lastCol As Integer
    lastCol = Source.Cells(1, Source.Columns.count).End(xlToLeft).Column
    Dim lastRow As Integer
    lastRow = Source.Cells(Source.Rows.count, 1).End(xlUp).Row
    
    'Get each trade row
    Dim tempArr() As Variant
    tempArr = Array(Source.Range(Source.Cells(2, 1), Source.Cells(lastRow, lastCol)).Value)
    TradeRows = tempArr(0)
End Sub

Sub GetAccounts()
    'Define a new collection
    Set AccountList = New Collection
    
    'Loop through each trade row element, get each account and their respective trades
    Dim ele As Integer
    For ele = 1 To UBound(TradeRows, 1)
        'Set the trade row as a new trade row object
        Dim tempRow As clsTradeRow
        Set tempRow = New clsTradeRow
        tempRow.Symbol = TradeRows(ele, GetIndexOf(HeaderStrings, "Symbol"))
        tempRow.Description = TradeRows(ele, GetIndexOf(HeaderStrings, "Description"))
        tempRow.Subclass = TradeRows(ele, GetIndexOf(HeaderStrings, "SubClass"))
        tempRow.Action = TradeRows(ele, GetIndexOf(HeaderStrings, "Action"))
        tempRow.Trade = TradeRows(ele, GetIndexOf(HeaderStrings, "Trade"))
        tempRow.Percent = TradeRows(ele, GetIndexOf(HeaderStrings, "PCNTSOLD"))
        
        'Loop through each account in the account list to see if it's already there
        Dim isAccountInList As Boolean
        isAccountInList = False
        Dim acct As Variant
        For Each acct In AccountList
            If Trim(TradeRows(ele, GetIndexOf(HeaderStrings, "AccountNumber"))) = acct.Number Then
                'Account is already in the list, add the trade row to the account
                acct.TradeList.Add tempRow
                
                isAccountInList = True
            End If
        Next acct
        
        If Not isAccountInList Then
            'The account isn't in the list, create a new account
            Dim tempAccount As clsAccount
            Set tempAccount = New clsAccount
            tempAccount.Number = TradeRows(ele, GetIndexOf(HeaderStrings, "AccountNumber"))
            tempAccount.Name = TradeRows(ele, GetIndexOf(HeaderStrings, "CRAccountMasterDescription"))
            tempAccount.AcctType = TradeRows(ele, GetIndexOf(HeaderStrings, "AccountType"))
            tempAccount.Custodian = TradeRows(ele, GetIndexOf(HeaderStrings, "Custodian"))
            
            'Add the trade row to a new account
            tempAccount.TradeList.Add tempRow
            
            'Add the new account to the list
            AccountList.Add tempAccount
        End If
    Next ele
End Sub

Sub ProcessAccount()
    'Replace cash with the appropriate money market
    CashToMM
    
    'Combine sells from different lots and delete extra rows
    CombineSameSymbols
    
    'Sort the account's transactions
    SortTrans
End Sub

Sub CashToMM()
    'If "CASH" in the list of transactions, combine it with the money market
    'Find if "CASH" is a symbol in the transaction list
    Dim cashIndex As Integer
    cashIndex = CashFoundAt
    If cashIndex <> -1 Then
        '"CASH" was found. Get the account's money market symbol
        Dim mmSymbol As String
        mmSymbol = GetMMSymbol
        If mmSymbol = vbNullString Then
            'There's only a cash transaction; no other money market. Change "CASH" to be the default
            CurrentAcct.TradeList(cashIndex).Symbol = "MMDA12"
            CurrentAcct.TradeList(cashIndex).Description = "TD BANK FDIC Insured Money Market"
        Else
            'Combine the cash transaction with the money market
            Dim symb As Variant
            For Each symb In CurrentAcct.TradeList
                If symb.Symbol = mmSymbol Then
                    'Add the cash trade to the money market trade
                    symb.Trade = symb.Trade + CurrentAcct.TradeList(cashIndex).Trade
                End If
            Next symb
        End If
        
        'Remove the cash transaction from the trade list
        CurrentAcct.TradeList.Remove cashIndex
    End If
End Sub

Function CashFoundAt() As Integer
    CashFoundAt = -1
    Dim trans As Integer
    For trans = 1 To CurrentAcct.TradeList.count
        If CurrentAcct.TradeList(trans).Symbol = "CASH" Then
            'Return the row the cash was found at
            CashFoundAt = trans
        End If
    Next trans
End Function

Function GetMMSymbol() As String
    'Return the symbol for the primary money market fund in the account
    Dim Trade As Variant
    For Each Trade In CurrentAcct.TradeList
        If Trade.Subclass = "MMM" And Trade.Symbol <> "CASH" Then
            GetMMSymbol = Trade.Symbol
        End If
    Next Trade
End Function

Sub CombineSameSymbols()
    'A security can have sales across many lots. Combine them all together
    'For each transaction, see if there is another transaction of the same symbol
    Dim trans As Integer
    trans = 1
    Do While trans <= CurrentAcct.TradeList.count
        Dim initialTrade As clsTradeRow
        Set initialTrade = CurrentAcct.TradeList(trans)
        'Go through transactions backwards so they can be removed if they have the same symbol
        Dim i As Integer
        For i = CurrentAcct.TradeList.count To trans + 1 Step -1
            Dim dupeTrade As clsTradeRow
            Set dupeTrade = CurrentAcct.TradeList(i)
            If initialTrade.Symbol = dupeTrade.Symbol Then
                'The transaction has the same symbol. Add the trade amount and percent sold to the first one
                initialTrade.Trade = initialTrade.Trade + dupeTrade.Trade
                initialTrade.Percent = initialTrade.Percent + dupeTrade.Percent
                
                'Remove the extra transaction
                CurrentAcct.TradeList.Remove i
                
                'If we're selling 100%, change the action to "SELL ALL"
                If initialTrade.Percent = 1 Then
                    initialTrade.Action = "SELL ALL"
                End If
            End If
        Next i
        trans = trans + 1
    Loop
End Sub

Sub SortTrans()
    'Create a temporary list to hold the sorted transactions
    Dim tempTrans As Collection
    Set tempTrans = New Collection
    
    'Fill the temporary list with the transactions, alphabetically by action and then by symbol
    Dim trans As Integer
    trans = CurrentAcct.TradeList.count
    Do While trans > 0
        Dim firstTrans As Integer
        firstTrans = trans
        
        'See if there's another transaction that should go before firstTrans
        Dim i As Integer
        For i = 1 To CurrentAcct.TradeList.count - 1
            If CurrentAcct.TradeList(i).Action > CurrentAcct.TradeList(firstTrans).Action Then
                firstTrans = i
            ElseIf CurrentAcct.TradeList(i).Action = CurrentAcct.TradeList(firstTrans).Action Then
                'If the actions are the same, take the transaction with the first alphabetical symbol
                If CurrentAcct.TradeList(i).Symbol < CurrentAcct.TradeList(firstTrans).Symbol Then
                    firstTrans = i
                End If
            End If
        Next i
        
        'Put the transaction into the temporary list and remove it from the account's transactions
        tempTrans.Add CurrentAcct.TradeList(firstTrans)
        CurrentAcct.TradeList.Remove firstTrans
        
        trans = trans - 1
    Loop
    
    'Put the sorted list back into the current account
    Set CurrentAcct.TradeList = tempTrans
End Sub

Sub GenerateReport()
    'Create a new worksheet for the report
    Set ClientSheet = clientBook.Worksheets.Add(Type:=xlWorksheet, After:=Source)
        
    'Put everything into the new sheet
    With ClientSheet
        'Add household name
        If household.Offset(1, 0).Value = "Schaefer, Russell & Patricia" Then
            'Client exception for the report
            .Range("A2").Value = "S, Russell & Patricia"
        Else
            'Put in the household name(s)
            .Range("A2").Value = household.Offset(1, 0).Value
        End If
        .Range("A2").Font.Bold = True
        .Range("A2:D2").Merge
        
        'Add the equity target
        .Range("E2").Value = "Equity Target"
        .Range("E2").Font.Underline = True
        .Range("E3").Value = eqTarget
        .Range("E3").HorizontalAlignment = xlLeft
        
        'Add the accounts and their transactions
        Dim pageFirstRow As Integer
        Dim acctStart As Range
        pageFirstRow = 1
        Set acctStart = .Range("A6")
        Dim acct As Variant
        For Each acct In AccountList
            Set CurrentAcct = acct
            AddAccountToReport acctStart, pageFirstRow
            
            'Go to next account
            Set acctStart = acctStart.Offset(CurrentAcct.TradeList.count + 7, 0)
        Next acct
    
        'Format sheet
        FormatClientSheet
    End With
End Sub

Sub AddAccountToReport(ByRef startRange As Range, ByRef pageFirstRow As Integer)
    'If the account's transactions would be on another page, create a page break and add a line on top of the new page
    If startRange.Row + CurrentAcct.TradeList.count + 6 - pageFirstRow > 43 Then
        ClientSheet.Rows(startRange.Row).PageBreak = xlPageBreakManual
        Range(startRange, startRange.Offset(0, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        Range(startRange, startRange.Offset(0, 5)).Borders(xlEdgeTop).Weight = xlMedium
        pageFirstRow = startRange.Row
        Set startRange = startRange.Offset(1, 0)
    End If
        
    'Add account information
    AccountInfo startRange
        
    'Add trade information
    TradeInfo startRange
End Sub

Sub AccountInfo(acctStart As Range)
    With acctStart
        .EntireRow.RowHeight = 15
        
        'Account name
        .Value = CurrentAcct.Name
        .Font.Bold = True
        
        'Small space under account name
        .Offset(1, 0).EntireRow.RowHeight = 3
        
        'Account information
        .Offset(2, 0).Value = "Custodian"
        .Offset(2, 0).Font.Underline = True
        .Offset(3, 0).Value = CurrentAcct.Custodian
        .Offset(2, 2).Value = "Account Type"
        .Offset(2, 2).Font.Underline = True
        .Offset(3, 2).Value = CurrentAcct.AcctType
        .Offset(3, 2).HorizontalAlignment = xlLeft
        
        Range(.Offset(2, 0), .Offset(2, 1)).Merge
        Range(.Offset(2, 2), .Offset(2, 3)).Merge
        Range(.Offset(3, 0), .Offset(3, 1)).Merge
        Range(.Offset(3, 2), .Offset(3, 3)).Merge
        Range(acctStart, .Offset(0, 4)).Merge
    End With
End Sub

Sub TradeInfo(acctStart As Range)
    'Start the trades after the account information
    Dim tradeStart As Range
    Set tradeStart = acctStart.Offset(5, 0)
    
    With tradeStart
        'Trade headers
        .Offset(0, 0).Value = "Action"
        .Offset(0, 1).Value = "Trade"
        .Offset(0, 2).Value = "Symbol"
        .Offset(0, 3).Value = "Description"
        .Offset(0, 0).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Offset(0, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        Range(.Offset(0, 0), .Offset(0, 2)).HorizontalAlignment = xlCenter
        Range(.Offset(0, 3), .Offset(0, 4)).Merge
        Range(.Offset(0, 0), .Offset(0, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            
        'Trades
        Dim trans As Integer
        For trans = 1 To CurrentAcct.TradeList.count
            'Get the current trade
            Dim currentTrade As clsTradeRow
            Set currentTrade = CurrentAcct.TradeList(trans)
            
            'Add Action to the report
            .Offset(trans, 0).Value = currentTrade.Action
            .Offset(trans, 0).HorizontalAlignment = xlCenter
            .Offset(trans, 0).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Trade to the report
            Dim tradeAmount As Single
            tradeAmount = currentTrade.Trade
            'If trade ends in 99.99, round it up to the nearest dollar
            If Right(tradeAmount, 3) = ".99" Then
                'Have tradeAmount be an int to prevent garbage after the decimal
                If tradeAmount < 0 Then
                    'tradeAmount rounds down to the 100
                    tradeAmount = Int(tradeAmount)
                Else
                    'tradeAmount rounds down to 99, so add 1
                    tradeAmount = Int(tradeAmount) + 1
                End If
            End If
            .Offset(trans, 1).Value = tradeAmount
            .Offset(trans, 1).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
            .Offset(trans, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Symbol to the report
            .Offset(trans, 2).Value = UCase(currentTrade.Symbol)
            .Offset(trans, 2).HorizontalAlignment = xlCenter
            .Offset(trans, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Description to the report
            .Offset(trans, 3).Value = currentTrade.Description
            Range(.Offset(trans, 3), .Offset(trans, 5)).Merge
        Next trans
        
        'Ensure row height of all trades is equal and add space between trades and next account
        Range(.Offset(-3, 0), .Offset(CurrentAcct.TradeList.count)).EntireRow.RowHeight = 15
        .Offset(1 + CurrentAcct.TradeList.count, 0).RowHeight = 30
    End With
End Sub

Sub FormatClientSheet()
    With ClientSheet
        'Set rows and columns
        SetRowColSize
        
        'Border at top of the page between content and header
        .Range("A1:F1").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A1:F1").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Name = "Client Copy"
        .UsedRange.Font.Size = 11
        .UsedRange.Font.Name = "Arial"
    End With
        
    'Turn off print communication for better performance
    If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
        Application.PrintCommunication = False
    End If
    
    'Format page layout
    AddHeader
    AddFooter
    FormatPrintArea
    
    'Turn print communication back on
    If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
        Application.PrintCommunication = True
    End If
End Sub

Sub SetRowColSize()
    'Set column sizes
    Dim Widths() As Variant
    Widths = Array(10, 13, 13, 24, 20, 13.5)
    Dim i As Integer
    For i = 0 To 5
        ClientSheet.Columns(i + 1).ColumnWidth = Widths(i)
    Next i

    'Set row sizes
    ClientSheet.Rows(1).RowHeight = 3
    For i = 2 To 6
        ClientSheet.Rows(i).RowHeight = 15
    Next i
End Sub

Sub AddHeader()
With ClientSheet.PageSetup
    If Time < TimeValue("15:00:00") Then
        .LeftHeader = "&L" & chr(10) & chr(10) & "&""Arial""&12&BTrade Recommendations - " & Date
    Else
        If Weekday(Date, vbSunday) = 6 Then
            .LeftHeader = "&L" & chr(10) & chr(10) & "&""Arial""&12&BTrade Recommendations - " & Date + 3
        Else
            .LeftHeader = "&L" & chr(10) & chr(10) & "&""Arial""&12&BTrade Recommendations - " & Date + 1
        End If
    End If

    Dim imagePath As String
    imagePath = "Z:\DO NOT MOVE FPIS-logo-final.jpg"
    
    If Dir(imagePath) = "" Then
        MsgBox "FPIS logo not found. Header missing logo."
    Else
        .RightHeader = "&R&g"
        .RightHeaderPicture.fileName = imagePath
        .RightHeaderPicture.LockAspectRatio = msoTrue
        .RightHeaderPicture.Height = Application.InchesToPoints(0.7)
    End If
End With
End Sub

Sub AddFooter()
With ClientSheet.PageSetup
    Dim footerFormat As String
    Dim footerStr As String
    footerFormat = "&""Arial""&9&I"
    footerStr = "The recommendations outlined above are estimates subject to market fluctuations and may not be " _
        & "traded in the exact dollar value indicated."
    
    .LeftFooter = footerFormat & footerStr
End With
End Sub

Sub FormatPrintArea()
With ClientSheet.PageSetup
    .TopMargin = Application.InchesToPoints(1.2)
    .BottomMargin = Application.InchesToPoints(0.75)
    .LeftMargin = Application.InchesToPoints(0.4)
    .RightMargin = Application.InchesToPoints(0.4)
    .HeaderMargin = Application.InchesToPoints(0.5)
    .FooterMargin = Application.InchesToPoints(0.3)

    .Orientation = xlPortrait
End With
End Sub

Sub PrintReport()
    'Enable screen and print report
    Both.UpdateScreen "On"
    ClientSheet.DisplayPageBreaks = False
    ClientSheet.Columns("A:F").PrintOut Copies:=2, Collate:=True, Preview:=True
End Sub

Sub SaveReport()
    'Fill array with names of months
    Dim months() As String
    months = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
    
    'Set the file name. Default file name is "[Month] [Year].xlsx"
    Dim saveFile As String
    saveFile = Trim(months(Month(Date) - 1) & " " & Year(Date))
    
    'Open save dialog in order to select save location
    Dim sfdSaveReport As FileDialog
    Set sfdSaveReport = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveReport.InitialFileName = saveDir & saveFile
    sfdSaveReport.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveReport.SelectedItems.count > 0 Then
        savePath = sfdSaveReport.SelectedItems.Item(1)
        clientBook.SaveAs fileName:=savePath, FileFormat:=xlWorkbookDefault, ReadOnlyRecommended:=False
    End If
End Sub

Function FindHeader(str As String) As Range
    'Call the actual FindHeader function
    Both.FindHeader (str)
End Function

Function GetIndexOf(arr As Variant, str As String) As Integer
    'Return the index of an array's element
    Dim i As Integer
    GetIndexOf = -1
    i = 1
    Do While GetIndexOf = -1 And i <= UBound(arr, 2)
        If arr(1, i) = str Then
            GetIndexOf = i
        End If
        i = i + 1
    Loop
    
    If GetIndexOf = -1 Then
        ShowError str & " not found on TradeRecommendationsExport sheet. Macro has been halted."
    End If
End Function

Sub ShowError(errMessage As String)
    MsgBox errMessage
    Both.UpdateScreen "On"
    End
End Sub
