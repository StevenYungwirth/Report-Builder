Attribute VB_Name = "SubclassReport"
Option Explicit
Public SCReport As Worksheet
Private AccountList As Collection
Private SubclassList As Collection
Private PrintRange As Range

Sub BuildSubclassReport(AccountList As Collection)
    'Get the list of accounts
    Set AccountList = Reports.AccountList
    
    'Add each subclass that has trades
    GenerateSubclassList AccountList
    
    'Create a new worksheet for the report after the client's trades worksheet
    Set SCReport = ReportAndLetter.ClientBook.Worksheets.Add(Type:=xlWorksheet, After:=ClientBook.Worksheets(2))
    SCReport.Name = "Trades by Subclass"
    
    'For each subclass, put in the trades
    FillReport
    
    'Adjust column sizes, add the header and footer
    FormatReport
End Sub

Sub GenerateSubclassList(AccountList As Collection)
    'Initialize the subclass list to have every subclss
    InitializeSubclassList
    
    Dim account As Variant
    For Each account In AccountList
        'Put each trade into their respective subclass list
        FillSubclassList account
    Next account
End Sub

Sub InitializeSubclassList()
    'Set the list of possible subclasses as they appear in TRX
    Dim trxSubclasses() As Variant
    trxSubclasses = Array("ULCV", "ULCB", "G", "ULCG", "UMCV", "UMCB", "UMCG", "USCV", "USCB", "USCG", _
    "RE", "S", "IE", "NUDS", "NUES", "B", "C&E", "FI", "FIO", "NUDB", "UHYB", "UIPB", "UTITB", "UTSTB", _
    "UTEB", "MMM", "O", "NC", "UNC")
    
    'Set the list of possible subclasses as they should appear on the report
    Dim subclasses() As Variant
    subclasses = Array("Large Value", "Large Blend", "Growth", "Large Growth", "Mid Value", "Mid Blend", "Mid Growth", _
    "Small Value", "Small Blend", "Small Growth", "Real Estate", "Specialty", "International Equities", _
    "Non-US Developed Stock", "Non-US Emerging Stock", "Balanced", "Cash & Equivalents", "Fixed Income", _
    "Fixed Income - Other", "Non-US Developed Bonds", "High Yield Bonds", "Inflation-Protected Bonds", _
    "Taxable Intermediate-Term Bonds", "Taxable Short-Term Bonds", "Tax-Exempt Bonds", "Money Market", _
    "Other", "Not Classified", "Unclassified")
        
    'Create a new clsSubclass for each subclass and add it to the list
    Set SubclassList = New Collection
    Dim tempSubclass As clsSubclass
    Dim sc As Integer
    For sc = 0 To UBound(trxSubclasses)
        Set tempSubclass = New clsSubclass
        tempSubclass.TRXDescription = trxSubclasses(sc)
        tempSubclass.Description = subclasses(sc)
        SubclassList.Add tempSubclass
    Next sc
End Sub

Sub FillSubclassList(account As Variant)
    Dim acctFund As Variant
    For Each acctFund In account.TradeList
        'Get the trade's subclass
        Dim acctFundSC As String
        acctFundSC = acctFund.Subclass
        
        'Find the respective subclass in the list
        Dim listSubclass As Variant
        Set listSubclass = GetSubclassFromList(acctFundSC)
        
        'Add the security as a trade in its subclass
        If listSubclass.TradeList.count = 0 Then
            'The trade list is empty, add the fund
            listSubclass.TradeList.Add acctFund
        Else
            'See if the security is already in the list
            Dim isInList As Boolean
            Dim scFund As Variant
            For Each scFund In listSubclass.TradeList
                If scFund.Symbol = acctFund.Symbol Then
                    'The security is already in the list, combine it with what's already there
                    scFund.Trade = scFund.Trade + acctFund.Trade
                    isInList = True
                End If
            Next scFund
            
            If Not isInList Then
                'The security wasn't already in the list, add it
                listSubclass.TradeList.Add acctFund
            End If
        End If
    Next acctFund
End Sub

Function GetSubclassFromList(sc As String) As clsSubclass
    Dim listSubclass As Variant
    For Each listSubclass In SubclassList
        If sc = listSubclass.TRXDescription Then
            Set GetSubclassFromList = listSubclass
        End If
    Next listSubclass
End Function

Sub FillReport()
    'Start on the second row to have a small space between the header and content
    Set PrintRange = SCReport.Range("A3")
    
    'For each subclass, if it has at least one trade, put it on the report
    Dim sc As Variant
    For Each sc In SubclassList
        If sc.TradeList.count > 0 Then
            PutSubclassOnReport sc
        End If
    Next sc
End Sub

Sub NextLine()
    Set PrintRange = PrintRange.Offset(1, 0)
End Sub

Sub PutSubclassOnReport(sc As Variant)
    'Get the starting position in case a page break needs to be added
    Dim startRange As Range
    Dim startPageBreaks As Integer
    Set startRange = PrintRange
    startPageBreaks = SCReport.HPageBreaks.count
    
    'Add the subclass description
    PrintRange.Value = sc.Description
    PrintRange.Font.Bold = True
    SCReport.Range(PrintRange, PrintRange.Offset(0, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    SCReport.Range(PrintRange, PrintRange.Offset(0, 2)).Borders(xlEdgeBottom).Weight = xlMedium
    
    'Go to the next line
    NextLine
    
    'There's at least one trade, put every trade in the subclass onto the report
    Dim total As Double
    Dim fund As Variant
    For Each fund In sc.TradeList
        total = total + fund.Trade
        PutFundOnReport fund
    Next fund
    
    'Put the total transaction amount by the subclass header
    startRange.Offset(0, 2).Value2 = total
    startRange.Offset(0, 2).Font.Bold = True
    startRange.Offset(0, 2).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
    
    'Add a space between subclasses
    NextLine
    
    'If the subclass is split between pages, put it all on the new page
    Dim endPageBreaks As Integer
    endPageBreaks = SCReport.HPageBreaks.count
    If startPageBreaks <> endPageBreaks Then
        'Add a page break before the subclass section
        startRange.Offset(-1, 0).EntireRow.insert
        SCReport.Rows(startRange.Offset(-2, 0).Row).PageBreak = xlPageBreakManual
        
        'Put a border at the top between the header and the content
        Range(startRange, startRange.Offset(-2, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        Range(startRange, startRange.Offset(-2, 5)).Borders(xlEdgeTop).Weight = xlMedium
    End If
End Sub

Sub PutFundOnReport(fund As Variant)
    'Fund name
    PrintRange.Value2 = fund.Description
    
    'Symbol
    PrintRange.Offset(0, 1).Value2 = fund.Symbol
    PrintRange.Offset(0, 1).HorizontalAlignment = xlCenter
    
    'Amount
    PrintRange.Offset(0, 2).Value2 = fund.Trade
    PrintRange.Offset(0, 2).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
    
    'Increment the print range
    NextLine
End Sub

Sub FormatReport()
    'Change the column sizes
    SCReport.Columns(1).ColumnWidth = 45
    SCReport.Columns(2).ColumnWidth = 13
    SCReport.Columns(3).ColumnWidth = 13
    SCReport.Columns(4).ColumnWidth = 8
    SCReport.Columns(5).ColumnWidth = 7
    SCReport.Columns(6).ColumnWidth = 7
    
    'Format the header, footer, and print area
    Reports.FormatClientSheet SCReport
End Sub
