Attribute VB_Name = "TradeReport"
Option Explicit
Public TrReport As Worksheet
Private CurrentAcct As clsAccount
Private AccountList As Collection

Sub BuildTradeReport()
    'Get the list of accounts
    Set AccountList = Reports.AccountList
    
    'Process each account
    Dim acct As Variant
    For Each acct In AccountList
        'Add any cash trades to the money market trades, combine sells from different trade lots, and sort the transactions
        Set CurrentAcct = acct
        ProcessAccount
    Next acct
    
    'Create a new worksheet for the report
    Set TrReport = ReportAndLetter.ClientBook.Worksheets.Add(Type:=xlWorksheet, After:=ReportAndLetter.ExportSheet)
    TrReport.Name = "Client Trades"
    
    'Generate the report
    GenerateReport
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
            End If
        Next i
        
        'If we're selling 100%, change the action to "SELL ALL"
        If initialTrade.Percent = 1 Then
            initialTrade.Action = "SELL ALL"
        End If
            
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
        
    'Put everything into the new sheet
    With TrReport
        'Get household name from Both
        Dim householdName As String
        householdName = ReportAndLetter.household.Offset(1, 0).Value
        
        If householdName = "Schaefer, Russell & Patricia" Then
            'Client exception for the report
            .Range("A2").Value = "S, Russell & Patricia"
        Else
            'Put in the household name(s)
            .Range("A2").Value = householdName
        End If
        .Range("A2").Font.Bold = True
        .Range("A2:D2").Merge
        
        'Add the equity target
        .Range("E2").Value = "Equity Target"
        .Range("E2").Font.Underline = True
        .Range("E3").Value = ReportAndLetter.eqTarget
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
    
        'Set rows and columns
        SetRowColSize
        
        'Format the header, footer, and print area
        Reports.FormatClientSheet TrReport
    End With
End Sub

Sub AddAccountToReport(ByRef startRange As Range, ByRef pageFirstRow As Integer)
    'If the account's transactions would be on another page, create a page break and add a line on top of the new page
    If startRange.Row + CurrentAcct.TradeList.count + 6 - pageFirstRow > 43 Then
        TrReport.Rows(startRange.Row).PageBreak = xlPageBreakManual
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
        .Offset(0, 0).Value2 = "Action"
        .Offset(0, 1).Value2 = "Trade"
        .Offset(0, 2).Value2 = "Symbol"
        .Offset(0, 3).Value2 = "Description"
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

Sub SetRowColSize()
    'Set column sizes
    Dim Widths() As Variant
    Widths = Array(10, 13, 13, 24, 20, 13.5)
    Dim i As Integer
    For i = 0 To 5
        TrReport.Columns(i + 1).ColumnWidth = Widths(i)
    Next i

    'Set row sizes
    TrReport.Rows(1).RowHeight = 3
    For i = 2 To 6
        TrReport.Rows(i).RowHeight = 15
    Next i
End Sub
