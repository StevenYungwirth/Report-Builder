Attribute VB_Name = "SubclassReport"
Option Explicit
Private PrintRange As Range

Public Function BuildSubclassReport(household As clsHousehold) As Worksheet
    'Add each subclass that has trades
    Dim subclassList As Collection
    Set subclassList = InitializeSubclassList(household.Accounts)
    
    'Create a new worksheet for the report after the client's trades worksheet
    With household.tradeSheet.Window.Book
        Dim scReport As Worksheet
        Set scReport = .Worksheets.Add(Type:=xlWorksheet, After:=.Worksheets(.Worksheets.count))
        scReport.Name = "Trades by Subclass"
    End With
    
    'For each subclass, put in the trades
    FillReport scReport, subclassList
    
    'Adjust column sizes, add the header and footer
    FormatReport scReport, household.Name, household.eqTarget
    
    Set BuildSubclassReport = scReport
End Function

Private Function InitializeSubclassList(accountList As Collection) As Collection
    'Set the list of possible subclasses as they appear in TRX
    Dim trxSubclasses() As String
    trxSubclasses = Split("ULCV,ULCB,G,ULCG,UMCV,UMCB,UMCG,USCV,USCB,USCG,RE,S,IE,NUDS,NUES,B,C&E," _
    & "FI,FIO,NUDB,UHYB,UIPB,UTITB,UTSTB,UTEB,MMM,O,NC,UNC", ",")
    
    'Set the list of possible subclasses as they should appear on the report
    Dim subclasses() As String
    subclasses = Split("Large Value,Large Blend,Growth,Large Growth,Mid Value,Mid Blend,Mid Growth," _
    & "Small Value,Small Blend,Small Growth,Real Estate,Specialty,International Equities,Non-US Developed Stock," _
    & "Non-US Emerging Stock,Balanced,Cash & Equivalents,Fixed Income,Fixed Income - Other,Non-US Developed Bonds," _
    & "High Yield Bonds,Inflation-Protected Bonds,Taxable Intermediate-Term Bonds,Taxable Short-Term Bonds," _
    & "Tax-Exempt Bonds,Money Market,Other,Not Classified,Unclassified", ",")
        
    'Create a new clsSubclass for each subclass and add it to the list
    Dim subclassList As Collection
    Set subclassList = New Collection
    Dim sc As Integer
    For sc = 0 To UBound(trxSubclasses)
        Dim tempSubclass As clsSubclass
        Set tempSubclass = ClassBuilder.NewSubclass(trxSubclasses(sc), subclasses(sc), accountList)
        subclassList.Add tempSubclass
    Next sc
    
    'Return the subclass list
    Set InitializeSubclassList = subclassList
End Function

Private Sub FillReport(scReport As Worksheet, subclassList As Collection)
    'Start on the second row to have a small space between the header and content
    Set PrintRange = scReport.Range("A6")
    
    'For each subclass, if it has at least one trade, put it on the report
    Dim sc As Variant
    For Each sc In subclassList
        If sc.TradeList.count > 0 Then
            PutSubclassOnReport scReport, sc
        End If
    Next sc
End Sub

Private Sub PutSubclassOnReport(report As Worksheet, sc As Variant)
    'Get the starting position in case a page break needs to be added later
    Dim startRange As Range
    Set startRange = PrintRange
    
    'Get the current number of page breaks
    Dim startPageBreaks As Integer
    startPageBreaks = report.HPageBreaks.count
    
    'Add the subclass description
    PrintRange.Value2 = sc.Description
    PrintRange.Font.Bold = True
    report.Range(PrintRange, PrintRange.Offset(0, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    report.Range(PrintRange, PrintRange.Offset(0, 2)).Borders(xlEdgeBottom).Weight = xlMedium
    
    'Go to the next line
    NextLine
    
    'Put every trade in the subclass onto the report, and keep the running total of trades
    Dim total As Double
    Dim fund As Variant
    For Each fund In sc.TradeList
        total = total + fund.Amount
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
    endPageBreaks = report.HPageBreaks.count
    If startPageBreaks <> endPageBreaks Then
        'Add a page break before the subclass section
        ReportProcedures.AddPageBreak report, startRange
    End If
End Sub

Private Sub PutFundOnReport(fund As Variant)
    'Fund name
    PrintRange.Value2 = fund.Description
    
    'Symbol
    PrintRange.Offset(0, 1).Value2 = fund.Symbol
    PrintRange.Offset(0, 1).HorizontalAlignment = xlCenter
    
    'Amount
    PrintRange.Offset(0, 2).Value2 = fund.Amount
    PrintRange.Offset(0, 2).NumberFormat = "$#,###.00_.;[Black]-$#,###.00_."
    
    'Increment the print range
    NextLine
End Sub

Private Sub NextLine()
    Set PrintRange = PrintRange.Offset(1, 0)
End Sub

Private Sub FormatReport(scReport As Worksheet, hhName As String, eqTarget As String)
    With scReport
        'Change the column sizes
        .Columns(1).ColumnWidth = 45
        .Columns(2).ColumnWidth = 13
        .Columns(3).ColumnWidth = 15
        .Columns(4).ColumnWidth = 13
        .Columns(5).ColumnWidth = 4
        .Columns(6).ColumnWidth = 3
    End With
    
    'Format the header, footer, and print area
    ReportProcedures.FormatClientSheet scReport, hhName, eqTarget
End Sub
