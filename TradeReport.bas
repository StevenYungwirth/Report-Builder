Attribute VB_Name = "TradeReport"
Option Explicit
Public TrReport As Worksheet
Private CurrentAcct As clsAccount
Private AccountList As Collection
Private PrintRange As Range

Sub BuildTradeReport()
    'Get the list of accounts
    Set AccountList = BothButton.AccountList
    
    'Process each account
    Dim acct As Variant
    For Each acct In AccountList
        'Add any cash trades to the money market trades, combine sells from different trade lots, and sort the transactions
        Set CurrentAcct = acct
        ProcessAccount
    Next acct
    
    'Create a new worksheet for the report
    Set TrReport = BothButton.ExportBook.Worksheets.Add(Type:=xlWorksheet, After:=BothButton.ExportBook.Worksheets(ExportBook.Worksheets.count))
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
        If Trade.subclass = "MMM" And Trade.Symbol <> "CASH" Then
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
                'The transaction has the same symbol. Add the trade amount to the first one
                initialTrade.Trade = initialTrade.Trade + dupeTrade.Trade
                
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
    'Set rows and columns
    SetRowColSize
    
    'Format the header, footer, and print area
    ReportsButton.FormatClientSheet TrReport
    
    'Put everything into the new sheet
    With TrReport
        'Add the accounts and their transactions
        Dim pageFirstRow As Integer
        pageFirstRow = 1
        Set PrintRange = .Range("A6")
        Dim acct As Variant
        For Each acct In AccountList
            Set CurrentAcct = acct
            AddAccountToReport pageFirstRow
    
            'Add a double space between trades and next account
            NextLine
            NextLine
        Next acct
    End With
End Sub

Sub AddAccountToReport(ByRef pageFirstRow As Integer)
    'Get the starting position in case a page break needs to be added
    Dim startRange As Range
    Dim startPageBreaks As Integer
    Set startRange = PrintRange
    startPageBreaks = TrReport.HPageBreaks.count
    
    'Put a value into the PrintRange cell to see if the new account starts on a new page
    PrintRange.Value2 = 0
    
    'If the account starts on a new page, put the page break where it should be and add a border at the top of it
    Dim preAddPageBreaks
    preAddPageBreaks = TrReport.HPageBreaks.count
    If startPageBreaks <> preAddPageBreaks Then
        'Add the page break with a space above the account information
        TrReport.Rows(PrintRange.Offset(-1, 0).Row).PageBreak = xlPageBreakManual
        
        'Put a border at the top between the header and the content
        TrReport.Range(PrintRange, PrintRange.Offset(-1, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        TrReport.Range(PrintRange, PrintRange.Offset(-1, 5)).Borders(xlEdgeTop).Weight = xlMedium
        
        'Set startPageBreaks to be the new count
        startPageBreaks = TrReport.HPageBreaks.count
    End If
        
    'Add account information
    AccountInfo
    
    'Add space between account information and trades
    NextLine
    NextLine
    
    'Add trade information
    TradeInfo
    
    'If the account is split between pages, put it all on the new page
    Dim endPageBreaks As Integer
    endPageBreaks = TrReport.HPageBreaks.count
    If startPageBreaks <> endPageBreaks Then
        'Add a page break before the account information
        TrReport.Rows(startRange.Offset(-1, 0).Row).PageBreak = xlPageBreakManual
        
        'Put a border at the top between the header and the content
        TrReport.Range(startRange, startRange.Offset(-1, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        TrReport.Range(startRange, startRange.Offset(-1, 5)).Borders(xlEdgeTop).Weight = xlMedium
    End If
End Sub

Sub AccountInfo()
    PrintRange.EntireRow.RowHeight = 15
    
    'Account name
    PrintRange.Value = CurrentAcct.Name
    PrintRange.Font.Bold = True
    Range(PrintRange, PrintRange.Offset(0, 4)).Merge
    NextLine
    
    
    'Small space under account name
    PrintRange.EntireRow.RowHeight = 3
    NextLine
    
    'Account information
    PrintRange.Value = "Custodian"
    PrintRange.Font.Underline = True
    PrintRange.Offset(0, 2).Value = "Account Type"
    PrintRange.Offset(0, 2).Font.Underline = True
    Range(PrintRange.Offset(0, 2), PrintRange.Offset(0, 3)).Merge
    Range(PrintRange, PrintRange.Offset(0, 1)).Merge
    NextLine
    
    PrintRange.Value = CurrentAcct.Custodian
    PrintRange.Offset(0, 2).Value = CurrentAcct.AcctType
    PrintRange.Offset(0, 2).HorizontalAlignment = xlLeft
    Range(PrintRange.Offset(0, 2), PrintRange.Offset(0, 3)).Merge
    Range(PrintRange, PrintRange.Offset(0, 1)).Merge
End Sub

Sub TradeInfo()
    'Trade headers
    PrintRange.Offset(0, 0).Value2 = "Action"
    PrintRange.Offset(0, 1).Value2 = "Trade"
    PrintRange.Offset(0, 2).Value2 = "Symbol"
    PrintRange.Offset(0, 3).Value2 = "Description"
    PrintRange.Offset(0, 0).Borders(xlEdgeRight).LineStyle = xlContinuous
    PrintRange.Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    PrintRange.Offset(0, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(PrintRange.Offset(0, 0), PrintRange.Offset(0, 2)).HorizontalAlignment = xlCenter
    Range(PrintRange.Offset(0, 3), PrintRange.Offset(0, 4)).Merge
    Range(PrintRange.Offset(0, 0), PrintRange.Offset(0, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    NextLine
    
    'Trades
    Dim trans As Integer
    For trans = 1 To CurrentAcct.TradeList.count
        'Get the current trade
        Dim currentTrade As clsTradeRow
        Set currentTrade = CurrentAcct.TradeList(trans)
            
        'Add Action to the report
        PrintRange.Value = currentTrade.Action
        PrintRange.HorizontalAlignment = xlCenter
        PrintRange.Borders(xlEdgeRight).LineStyle = xlContinuous
        
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
        PrintRange.Offset(0, 1).Value = tradeAmount
        PrintRange.Offset(0, 1).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
        PrintRange.Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        
        'Add Symbol to the report
        PrintRange.Offset(0, 2).Value = currentTrade.Symbol
        PrintRange.Offset(0, 2).HorizontalAlignment = xlCenter
        PrintRange.Offset(0, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        
        'Add Description to the report
        PrintRange.Offset(0, 3).Value = currentTrade.Description
        Range(PrintRange.Offset(0, 3), PrintRange.Offset(0, 5)).Merge
        
        'Ensure row height of the trade is correct and go to the next trade
        PrintRange.Rows.EntireRow.RowHeight = 15
        NextLine
    Next trans
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
    For i = 2 To 6
        TrReport.Rows(i).RowHeight = 15
    Next i
End Sub

Sub NextLine()
    Set PrintRange = PrintRange.Offset(1, 0)
End Sub
