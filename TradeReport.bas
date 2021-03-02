Attribute VB_Name = "TradeReport"
Option Explicit
Private PrintRange As Range

Public Function BuildTradeReport(household As clsHousehold) As Worksheet
    'Create a new worksheet for the report
    With household.tradeSheet.Window.Book
        Dim TrReport As Worksheet
        Set TrReport = .Worksheets.Add(Type:=xlWorksheet, After:=.Worksheets(.Worksheets.count))
    End With
    TrReport.Name = "Client Trades"
    
    'Generate the report
    GenerateReport TrReport, household
    
    Set BuildTradeReport = TrReport
End Function

Private Sub GenerateReport(report As Worksheet, household As clsHousehold)
    'Set rows and columns
    SetRowColSize report
    
    'Format the header, footer, and print area
    ReportProcedures.FormatClientSheet sht:=report, hhName:=household.Name, eqTarget:=household.eqTarget
    
    'Put everything into the new sheet
    'Set the range where the report will start
    Set PrintRange = report.Range("A6")
    
    'Add the accounts and their transactions
    Dim acct As Variant
    For Each acct In household.Accounts
        AddAccountToReport acct, report

        'Add a double space between trades and next account
        NextLine
        NextLine
        
    Next acct
End Sub

Private Sub SetRowColSize(report As Worksheet)
    'Set column sizes
    Dim Widths() As Variant
    Widths = Array(10, 13, 13, 24, 20, 13.5)
    Dim i As Integer
    For i = 0 To 5
        report.Columns(i + 1).ColumnWidth = Widths(i)
    Next i

    'Set row sizes
    For i = 2 To 6
        report.Rows(i).RowHeight = 15
    Next i
End Sub

Private Sub AddAccountToReport(acct As Variant, report As Worksheet)
    'Get the starting position in case a page break needs to be added
    Dim startRange As Range
    Set startRange = PrintRange
    
    'Get the current number of page breaks
    Dim startPageBreaks As Integer
    startPageBreaks = report.HPageBreaks.count
    
    'Put a test value into the PrintRange cell to see if acct would start on a new page
    PrintRange.Value2 = 0
    
    'If the account starts on a new page, put the page break where it should be and add a border at the top of it
    Dim preAddPageBreaks As Integer
    preAddPageBreaks = report.HPageBreaks.count
    If startPageBreaks <> preAddPageBreaks Then
        ReportProcedures.AddPageBreak report, PrintRange
        
        'Set startPageBreaks to be the new count
        startPageBreaks = report.HPageBreaks.count
    End If
        
    'Add account information
    AddAccountInfo acct
    
    'Add space between account information and trades
    NextLine
    NextLine
    
    'Add trade information
    AddTradeInfo acct
    
    'If the account is split between pages, put it all on the new page
    Dim endPageBreaks As Integer
    endPageBreaks = report.HPageBreaks.count
    If startPageBreaks <> endPageBreaks Then
        ReportProcedures.AddPageBreak report, startRange
    End If
End Sub

Private Sub AddAccountInfo(acct As Variant)
    PrintRange.EntireRow.RowHeight = 15
        
    With PrintRange
        'Account name
        .value = acct.Name
        .Font.Bold = True
        Range(.Offset, .Offset(0, 4)).Merge
        NextLine
    End With
        
    With PrintRange
        'Small space under account name
        .EntireRow.RowHeight = 3
        NextLine
    End With
        
    With PrintRange
        'Account information
        .value = "Custodian"
        .Font.Underline = True
        .Offset(0, 2).value = "Account Type"
        .Offset(0, 2).Font.Underline = True
        Range(.Offset(0, 2), .Offset(0, 3)).Merge
        Range(.Offset, .Offset(0, 1)).Merge
        NextLine
    End With
        
    With PrintRange
        .value = acct.Custodian
        .Offset(0, 2).value = acct.acctType
        .Offset(0, 2).HorizontalAlignment = xlLeft
        Range(.Offset(0, 2), .Offset(0, 3)).Merge
        Range(.Offset, .Offset(0, 1)).Merge
    End With
End Sub

Private Sub AddTradeInfo(acct As Variant)
    With PrintRange
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
        NextLine
    End With
        
    'Trades
    Dim trans As Integer
    For trans = 1 To acct.TradeList.count
        With PrintRange
            'Get the current trade
            Dim currentTrade As clsTradeRow
            Set currentTrade = acct.TradeList(trans)
                
            'Add Action to the report
            .value = currentTrade.Action
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Trade to the report
            Dim tradeAmount As Single
            tradeAmount = currentTrade.Amount
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
            .Offset(0, 1).value = tradeAmount
            .Offset(0, 1).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
            .Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Symbol to the report
            .Offset(0, 2).value = currentTrade.Symbol
            .Offset(0, 2).HorizontalAlignment = xlCenter
            .Offset(0, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Add Description to the report
            .Offset(0, 3).value = currentTrade.Description
            Range(.Offset(0, 3), .Offset(0, 5)).Merge
            
            'Ensure row height of the trade is correct and go to the next trade
            .Rows.EntireRow.RowHeight = 15
            NextLine
        End With
    Next trans
End Sub

Private Sub NextLine()
    Set PrintRange = PrintRange.Offset(1, 0)
End Sub
