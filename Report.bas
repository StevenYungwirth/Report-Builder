Attribute VB_Name = "Report"
Option Explicit
Private ClientSheet As Worksheet
Private acctNumLoc As Range
Private symbol As Range
'Arrays holding which rows accounts start on and how many trades are in each account
Private AcctRows() As Variant
Private NumTrades() As Variant
Sub BuildReport(Optional bothClicked As Boolean)
    'Take exported csv from TRX and build the client's report from it
    'Set global variables
    Both.UpdateScreen "Off"
    On Error GoTo MacroBroke
    If Not bothClicked Then
        Both.SetGlobals
    End If
    SetReportGlobals
    
    'Replace cash with the appropriate money market
    CashToMM

    'Sort transactions by account
    SortTrans

    'Combine sells from different lots and delete extra rows
    FindExtra

    'Find rows where other accounts start, and get how many transactions are in each account
    AcctLocations

    'Create new sheet
    ClientCopy

    'Print Report
    PrintReport
    
    'Save the exported csv as an excel workbook in the client's folder
    SaveReport
    
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
    'Check the column headers of headers only used once or twice. If they don't work, the macro will end before anything happens
    FindHeader ("AccountNumber")
    FindHeader ("CRAccountMasterDescription")
    FindHeader ("Custodian")
    FindHeader ("Symbol")
    FindHeader ("OriginalTradeDate")
    FindHeader ("CostBasis")
    FindHeader ("Trade")
    FindHeader ("AccountType")
    FindHeader ("Action")
    FindHeader ("Description")
    FindHeader ("PCNTSOLD")

    'Set the column headers
    Set acctNumLoc = FindHeader("AccountNumber")
    Set symbol = FindHeader("Symbol")
End Sub
Sub CashToMM()
    Do While Not Source.Range(symbol, symbol.Offset(Source.UsedRange.Rows.count - 1)).Find("CASH") Is Nothing
        'Find location of the CASH symbol and its account number
        Dim cashRow As Integer
        Dim cashAcct As String
        cashRow = Source.Range(symbol, symbol.Offset(Source.UsedRange.Rows.count - 1)).Find("CASH").Row
        cashAcct = Source.Cells(cashRow, acctNumLoc.Column).Value
    
        'Get the bounds of the account
        Dim i As Integer
        Dim firstRow As Integer
        Dim lastRow As Integer
        firstRow = 0
        lastRow = 0
        For i = 2 To Source.UsedRange.Rows.count
            If Source.Cells(i, acctNumLoc.Column).Value = cashAcct Then
                If firstRow = 0 Then
                    firstRow = i
                Else
                    lastRow = i
                End If
            End If
        Next i
    
        'Default money market symbol and description
        Dim mmSym As String
        Dim mmDes As String
        mmSym = "MMDA12"
        mmDes = "TD Bank FDIC Insured Money Market"
        
        'Lists of possible money market symbol and descriptions
        Dim mmSymArr() As Variant
        Dim mmDesArr() As Variant
        mmDesArr = Array("TD Bank FDIC Insured Money Market", "FDIC Insured Deposit Account Core Not Covered By S", "FDIC Insured Deposit Account IDA02 Not Covered By", "FDIC Insured Deposit Account IDA09 Not Covered By")
        mmSymArr = Array("MMDA12", "MMDA1", "MMDA2", "ZFD90")
        
        'For each row in the account, if it contains a money market symbol, that's the one to use in place of CASH
        Dim j As Integer
        Dim k As Integer
        For j = 0 To UBound(mmSymArr)
            For k = firstRow To lastRow
                If Source.Cells(k, symbol.Column).Value = mmSymArr(j) Then
                    mmSym = mmSymArr(j)
                    mmDes = mmDesArr(j)
                End If
            Next k
        Next j
    
        Dim description As Range
        Set description = FindHeader("Description")
        Source.Cells(cashRow, symbol.Column).Value = mmSym
        Source.Cells(cashRow, description.Column).Value = mmDes
    Loop
End Sub
Sub SortTrans()
    Dim acctName As Range
    Set acctName = FindHeader("CRAccountMasterDescription")
    Dim action As Range
    Set action = FindHeader("Action")
    
    Dim firstCell As Range
    Dim lastCell As Range
    Set firstCell = Source.Cells(2, 1)
    Set lastCell = Source.Cells(Source.UsedRange.Rows.count, Source.UsedRange.Columns.count)
    
    Source.Range(firstCell, lastCell).Sort Key1:=symbol, order1:=xlAscending, Header:=xlNo
    Source.Range(firstCell, lastCell).Sort Key1:=action, order1:=xlDescending, Header:=xlNo
    Source.Range(firstCell, lastCell).Sort Key1:=acctNumLoc, order1:=xlDescending, Header:=xlNo
    Source.Range(firstCell, lastCell).Sort Key1:=acctName, order1:=xlAscending, Header:=xlNo
End Sub
Sub FindExtra()
'If the same security is being bought/sold in the same account, it is extra
Dim totalTransactions As Integer
Dim numExtraRows As Integer
Dim ExtraRows() As Integer
Dim transaction As Integer

ReDim ExtraRows(1 To 1) As Integer
totalTransactions = Source.UsedRange.Rows.count - 1
numExtraRows = 0

'Find which rows are extra
For transaction = 1 To totalTransactions
    'If the symbol and account number of a transaction equals the next transaction's symbol and account number, then add it to an array
    'Array holds which row numbers are extra
    If (symbol.Offset(transaction).Value = symbol.Offset(transaction + 1).Value _
    Or symbol.Offset(transaction).Value = "MMDA12" And symbol.Offset(transaction + 1).Value = "CASH" _
    Or symbol.Offset(transaction).Value = "CASH" And symbol.Offset(transaction + 1).Value = "MMDA12") _
    And acctNumLoc.Offset(transaction).Value = acctNumLoc.Offset(transaction + 1).Value Then
        numExtraRows = numExtraRows + 1
        ReDim Preserve ExtraRows(1 To numExtraRows)
        ExtraRows(numExtraRows) = transaction + 2 'Add one for the header, add one since the next row is the extra one
    End If
Next transaction

'If there are any extra rows, then delete them
If ExtraRows(1) > 0 Then
    DeleteExtra ExtraRows
End If

'If we're selling 100%, change the action to sell all
Dim action As Range
Dim percent As Range
Set action = FindHeader("Action")
Set percent = FindHeader("PCNTSOLD")
For transaction = 1 To totalTransactions
    If percent.Offset(transaction, 0) = 1 Then
        action.Offset(transaction, 0).Value = "SELL ALL"
    End If
Next transaction

'Sort the transactions again, in case a sell all was added
SortTrans
End Sub
Sub DeleteExtra(ExtraRows() As Integer)
'For each extra transaction, add the extra row's trade and cost basis values to the first row's trade and cost basis values, and then delete the extra row
    Dim trade As Range
    Dim tradeDate As Range
    Dim costBasis As Range
    Dim numExtraRows As Integer
    Dim extraTrans As Integer
    Dim firstTrans As Integer
    Dim firstTransValue As Single
    Dim extraTransValue As Single
    Dim totalTransValue
    
    Set trade = FindHeader("Trade")
    Set tradeDate = FindHeader("OriginalTradeDate")
    Set costBasis = FindHeader("CostBasis")
    numExtraRows = UBound(ExtraRows)
    For extraTrans = numExtraRows To 1 Step -1
        firstTrans = ExtraRows(extraTrans) - 1
        firstTransValue = Source.Cells(firstTrans, trade.Column).Value
        extraTransValue = Source.Cells(ExtraRows(extraTrans), trade.Column).Value
        totalTransValue = firstTransValue + extraTransValue
        
        'Add trade values together and cost basis together
        Source.Cells(firstTrans, trade.Column).Value = totalTransValue
        Source.Cells(firstTrans, costBasis.Column).Value = Source.Cells(firstTrans, costBasis.Column).Value + Source.Cells(firstTrans + 1, costBasis.Column).Value
        Source.Cells(firstTrans, tradeDate.Column).Value = "Multiple"
        
        'If the cash and money market are being combined, make sure the symbol is "MMDA12"
        If Source.Cells(firstTrans, symbol.Column).Value = "CASH" Then
            Source.Cells(firstTrans, symbol.Column).Value = "MMDA12"
        End If
        
        'Delete the extra row
        Rows(ExtraRows(extraTrans)).EntireRow.Delete
    Next extraTrans
End Sub
Sub AcctLocations()
'Find which rows different accounts start on
ReDim AcctRows(1 To 1) As Variant
ReDim NumTrades(1 To 1) As Variant
AcctRows(1) = acctNumLoc.Offset(1, 0).Row

Dim acctNum As Variant
Dim totalTransactions As Integer
Dim numOfAccts As Integer

acctNum = acctNumLoc.Offset(1, 0)
totalTransactions = Source.UsedRange.Rows.count - 1
numOfAccts = 1

'For all transactions, if the account number is different than the transaction before, add it to the array
Dim transaction As Integer
Dim account As Integer
For transaction = 1 To totalTransactions
    If acctNumLoc.Offset(transaction, 0) <> acctNum Then
        'Set the row the new account starts on, and subtract it from the previous row's start to find the number of transactions
        numOfAccts = numOfAccts + 1
        ReDim Preserve AcctRows(1 To numOfAccts)
        AcctRows(numOfAccts) = acctNumLoc.Offset(transaction, 0).Row
        
        ReDim Preserve NumTrades(1 To numOfAccts)
        NumTrades(numOfAccts - 1) = AcctRows(numOfAccts) - AcctRows(numOfAccts - 1)
        
        'Reset acctNum to be the new account number
        acctNum = acctNumLoc.Offset(transaction, 0)
    End If
Next transaction

'The number of trades in the last account
NumTrades(UBound(NumTrades)) = Source.UsedRange.Rows.count - AcctRows(numOfAccts) + 1
End Sub
Sub ClientCopy()
    Set ClientSheet = clientBook.Worksheets.Add(Type:=xlWorksheet, After:=Source)
        
    'Put everything into the new sheet
    With ClientSheet
        'Add household name, equity target
        If household.Offset(1, 0).Value = "Schaefer, Russell & Patricia" Then
            .Range("A2").Value = "S, Russell & Patricia"
        Else
            .Range("A2").Value = household.Offset(1, 0).Value
        End If
        .Range("E2").Value = "Equity Target"
        .Range("E3").Value = eqTarget
        .Range("A2").Font.Bold = True
        .Range("E2").Font.Underline = True
        .Range("E3").HorizontalAlignment = xlLeft
        .Range("A2:D2").Merge
        
        'Add the accounts with transcations and their transactions
        Dim pageFirstRow As Integer
        Dim acctStart As Range
        
        pageFirstRow = 1
        Set acctStart = .Range("A6")
    
        Dim account As Integer
        For account = 1 To UBound(AcctRows)
            'If the account's transactions would be on another page, create a page break and add a line on top of the new page
            If acctStart.Row + NumTrades(account) + 6 - pageFirstRow > 43 Then
                ClientSheet.Rows(acctStart.Row).PageBreak = xlPageBreakManual
                Range(acctStart, acctStart.Offset(0, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
                Range(acctStart, acctStart.Offset(0, 5)).Borders(xlEdgeTop).Weight = xlMedium
                pageFirstRow = acctStart.Row
                Set acctStart = acctStart.Offset(1, 0)
            End If
                
            'Add account information
            AccountInfo acctStart, account
                
            'Add trade information
            TradeInfo acctStart, account
            
            'Go to next account
            Set acctStart = acctStart.Offset(NumTrades(account) + 7, 0)
        Next account
        
        'Format sheet
        FormatClientSheet
    End With
End Sub
Sub AccountInfo(acctStart As Range, account As Integer)
    Dim acctName As Range
    Dim custodian As Range
    Dim acctType As Range
    
    Set acctName = FindHeader("CRAccountMasterDescription")
    Set custodian = FindHeader("Custodian")
    Set acctType = FindHeader("AccountType")
    
    With acctStart
        .EntireRow.RowHeight = 15
        
        'Account name
        .Value = acctName.Offset(AcctRows(account) - 1, 0).Value
        .Font.Bold = True
        
        'Small space under account name
        .Offset(1, 0).EntireRow.RowHeight = 3
        
        'Account information
        .Offset(2, 0).Value = "Custodian"
        .Offset(2, 2).Value = "Account Type"
        .Offset(2, 0).Font.Underline = True
        .Offset(2, 2).Font.Underline = True
        .Offset(3, 0).Value = custodian.Offset(AcctRows(account) - 1, 0).Value
        .Offset(3, 2).Value = acctType.Offset(AcctRows(account) - 1, 0).Value
        .Offset(3, 2).HorizontalAlignment = xlLeft
        
        Range(.Offset(2, 0), .Offset(2, 1)).Merge
        Range(.Offset(2, 2), .Offset(2, 3)).Merge
        Range(.Offset(3, 0), .Offset(3, 1)).Merge
        Range(.Offset(3, 2), .Offset(3, 3)).Merge
        Range(acctStart, .Offset(0, 4)).Merge
    End With
End Sub
Sub TradeInfo(acctStart As Range, account As Integer)
    Dim trade As Range
    Dim description As Range
    Dim action As Range
    Dim percent As Range
    Dim tradeStart As Range
    
    Set trade = FindHeader("Trade")
    Set description = FindHeader("Description")
    Set action = FindHeader("Action")
    Set percent = FindHeader("PCNTSOLD")
    Set tradeStart = acctStart.Offset(5, 0)
    
    With tradeStart
        'Trade headers
        .Value = "Action"
        .Offset(0, 1).Value = "Trade"
        .Offset(0, 2).Value = "Symbol"
        .Offset(0, 3).Value = "Description"
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Offset(0, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        Range(.Offset(0, 0), .Offset(0, 2)).HorizontalAlignment = xlCenter
        Range(.Offset(0, 3), .Offset(0, 4)).Merge
        Range(.Offset(0, 0), .Offset(0, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            
        'Trades
        Dim j As Integer
        For j = 1 To NumTrades(account)
            'Add Action to the report
            .Offset(j, 0).Value = action.Offset(AcctRows(account) + j - 2, 0).Value
            .Offset(j, 0).HorizontalAlignment = xlCenter
            .Offset(j, 0).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Trade to the report
            Dim tradeAmount As Single
            tradeAmount = trade.Offset(AcctRows(account) + j - 2, 0).Value
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
            .Offset(j, 1).Value = tradeAmount
            .Offset(j, 1).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
            .Offset(j, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Symbol to the report
            .Offset(j, 2).Value = symbol.Offset(AcctRows(account) + j - 2, 0).Value
            .Offset(j, 2).HorizontalAlignment = xlCenter
            .Offset(j, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Description to the report
            .Offset(j, 3).Value = description.Offset(AcctRows(account) + j - 2, 0).Value
            Range(.Offset(j, 3), .Offset(j, 5)).Merge
        Next j
        
        'Ensure row height of all trades is equal and add space between trades and next account
        Range(.Offset(-3, 0), .Offset(NumTrades(account))).EntireRow.RowHeight = 15
        .Offset(1 + NumTrades(account), 0).RowHeight = 30
    End With
End Sub
Sub FormatClientSheet()
    With ClientSheet
        'Set rows and columns
        RowColSize
        
        'Border at top of the page between content and header
        .Range("A1:F1").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A1:F1").Borders(xlEdgeTop).LineStyle = xlContinuous
        .name = "Client Copy"
        .UsedRange.Font.Size = 11
        .UsedRange.Font.name = "Arial"
        
        'Format page layout
        With .PageSetup
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
            
            Dim footerFormat As String
            Dim footerStr As String
            footerFormat = "&""Arial""&9&I"
            footerStr = "The recommendations outlined above are estimates subject to market fluctuations and may not be " _
                & "traded in the exact dollar value indicated."
            
            .LeftFooter = footerFormat & footerStr

            If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
                Application.PrintCommunication = False
            End If
            .TopMargin = Application.InchesToPoints(1.2)
            .BottomMargin = Application.InchesToPoints(0.75)
            .LeftMargin = Application.InchesToPoints(0.4)
            .RightMargin = Application.InchesToPoints(0.4)
            .HeaderMargin = Application.InchesToPoints(0.5)
            .FooterMargin = Application.InchesToPoints(0.3)

            .Orientation = xlPortrait
            
            If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
                Application.PrintCommunication = True
            End If
        End With
    End With
End Sub
Sub RowColSize()
    Dim i As Integer
    
    'Set column sizes
    Dim Widths() As Variant
    Widths = Array(10, 13, 13, 24, 20, 13.5)
    For i = 0 To 5
        ClientSheet.Columns(i + 1).ColumnWidth = Widths(i)
    Next i

    'Set row sizes
    ClientSheet.Rows(1).RowHeight = 3
    For i = 2 To 6
        ClientSheet.Rows(i).RowHeight = 15
    Next i
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
    
    'Default file name is [Month] [Year].xlsx
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
