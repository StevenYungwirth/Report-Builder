Attribute VB_Name = "Reports"
Option Explicit
Public AccountList As Collection

Sub GenerateReports(Optional bothClicked As Boolean)
    'Turn off screen updating to improve performance
    ReportAndLetter.UpdateScreen "Off"
    
    'Add vague error tracking
    On Error GoTo MacroBroke
    
    'Set global variables
    If Not bothClicked Then
        'Just the report is being generated, not both report and letter. Set the global variables
        '(TradeRecommendationsExport workbook and worksheet, household name, equity target, save location)
        ReportAndLetter.SetGlobals
    End If
    SetReportGlobals
    
    TradeReport.BuildTradeReport
    
    'Get household name from Both
    Dim householdName As String
    householdName = ReportAndLetter.household.Offset(1, 0).Value

    'Per client's request, include a report showing what's being bought/sold from each subclass
    Dim scReportRan As Boolean
    If householdName = "Herald, Robert & Catherine" Then
        SubclassReport.BuildSubclassReport AccountList
        scReportRan = True
    End If

    'Print the report(s)
    PrintReport TradeReport.TrReport
    If scReportRan Then
        PrintReport SubclassReport.SCReport
    End If
    
    'Save the exported csv as an excel workbook in the client's folder
    SaveReportWorkbook
    
    'Return everything back to normal
    Set TradeReport.TrReport = Nothing
    Set SubclassReport.SCReport = Nothing
    If Not bothClicked Then
        ReportAndLetter.ResetGlobals
    End If
    ReportAndLetter.UpdateScreen "On"
    Exit Sub
MacroBroke:
    ReportAndLetter.UpdateScreen "On"
    MsgBox "Fatal error, macro has halted"
End Sub

Sub SetReportGlobals()
    'Get the column headers
    Dim headerStrings() As Variant
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
    lastCol = ReportAndLetter.ExportSheet.Cells(1, ReportAndLetter.ExportSheet.Columns.count).End(xlToLeft).Column
    
    'Set the array to be the values of the first row
    Dim tempArr() As Variant
    Dim headers() As Variant
    tempArr = Array(ReportAndLetter.ExportSheet.Range(ReportAndLetter.ExportSheet.Cells(1, 1), ReportAndLetter.ExportSheet.Cells(1, lastCol)).Value)
    headers = tempArr(0)
    
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
    lastCol = ReportAndLetter.ExportSheet.Cells(1, ReportAndLetter.ExportSheet.Columns.count).End(xlToLeft).Column
    Dim lastRow As Integer
    lastRow = ReportAndLetter.ExportSheet.Cells(ReportAndLetter.ExportSheet.Rows.count, 1).End(xlUp).Row
    
    'Get each trade row
    Dim tempArr() As Variant
    tempArr = Array(ReportAndLetter.ExportSheet.Range(ReportAndLetter.ExportSheet.Cells(2, 1), ReportAndLetter.ExportSheet.Cells(lastRow, lastCol)).Value)
    GetTrades = tempArr(0)
End Function

Sub GetAccounts(headerStrings As Variant, tradeRows As Variant)
    'Define a new collection
    Set AccountList = New Collection
    
    'Loop through each trade row element, get each account and their respective trades
    Dim ele As Integer
    For ele = 1 To UBound(tradeRows, 1)
        'Set the trade row as a new trade row object
        Dim tempRow As clsTradeRow
        Set tempRow = New clsTradeRow
        tempRow.Symbol = tradeRows(ele, GetIndexOf(headerStrings, "Symbol"))
        tempRow.Description = tradeRows(ele, GetIndexOf(headerStrings, "Description"))
        tempRow.Subclass = tradeRows(ele, GetIndexOf(headerStrings, "SubClass"))
        tempRow.Action = tradeRows(ele, GetIndexOf(headerStrings, "Action"))
        tempRow.Trade = tradeRows(ele, GetIndexOf(headerStrings, "Trade"))
        tempRow.Percent = tradeRows(ele, GetIndexOf(headerStrings, "PCNTSOLD"))
        
        'Loop through each account in the account list to see if it's already there
        Dim isAccountInList As Boolean
        isAccountInList = False
        Dim acct As Variant
        For Each acct In AccountList
            If Trim(tradeRows(ele, GetIndexOf(headerStrings, "AccountNumber"))) = acct.Number Then
                'Account is already in the list, add the trade row to the account
                acct.TradeList.Add tempRow
                
                isAccountInList = True
            End If
        Next acct
        
        If Not isAccountInList Then
            'The account isn't in the list, create a new account
            Dim tempAccount As clsAccount
            Set tempAccount = New clsAccount
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

Sub FormatClientSheet(sht As Worksheet)
    With sht
        'Border at top of the page between content and header
        .Range("A1:F1").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A1:F1").Borders(xlEdgeTop).LineStyle = xlContinuous
        .UsedRange.Font.Size = 11
        .UsedRange.Font.Name = "Arial"
    End With
    
    'Format page layout
    AddHeader sht
    AddFooter sht
    FormatPrintArea sht
End Sub

Sub AddHeader(sht As Worksheet)
With sht.PageSetup
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

Sub AddFooter(sht As Worksheet)
With sht.PageSetup
    Dim footerFormat As String
    Dim footerStr As String
    footerFormat = "&""Arial""&9&I"
    footerStr = "The recommendations outlined above are estimates subject to market fluctuations and may not be " _
        & "traded in the exact dollar value indicated."
    
    .LeftFooter = footerFormat & footerStr
End With
End Sub

Sub FormatPrintArea(sht As Worksheet)
'Turn off print communication for better performance.
'This could make the header and footer faster as well, but they then don't fill properly for unknown reasons
If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
    Application.PrintCommunication = False
End If

With sht.PageSetup
    .TopMargin = Application.InchesToPoints(1.2)
    .BottomMargin = Application.InchesToPoints(0.75)
    .LeftMargin = Application.InchesToPoints(0.4)
    .RightMargin = Application.InchesToPoints(0.4)
    .HeaderMargin = Application.InchesToPoints(0.5)
    .FooterMargin = Application.InchesToPoints(0.3)

    .Orientation = xlPortrait
End With

'Turn print communication back on
If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
    Application.PrintCommunication = True
End If
End Sub

Sub PrintReport(sht As Worksheet)
    'Enable screen and print report
    ReportAndLetter.UpdateScreen "On"
    sht.DisplayPageBreaks = False
    sht.Columns("A:F").PrintOut Copies:=2, Collate:=True, Preview:=True
End Sub

Sub SaveReportWorkbook()
    'Fill array with names of months
    Dim months() As String
    months = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
    
    'Set the file name. Default file name is "[Month] [Year].xlsx"
    Dim saveFile As String
    saveFile = Trim(months(Month(Date) - 1) & " " & Year(Date))
    
    'Open save dialog in order to select save location
    Dim sfdSaveReport As FileDialog
    Set sfdSaveReport = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveReport.InitialFileName = ReportAndLetter.saveDir & saveFile
    sfdSaveReport.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveReport.SelectedItems.count > 0 Then
        savePath = sfdSaveReport.SelectedItems.Item(1)
        ReportAndLetter.ClientBook.SaveAs fileName:=savePath, FileFormat:=xlWorkbookDefault, ReadOnlyRecommended:=False
    End If
End Sub

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

Function FindHeader(str As String) As Range
    'Call the actual FindHeader function
    ReportAndLetter.FindHeader (str)
End Function

Sub ShowError(errMessage As String)
    MsgBox errMessage
    ReportAndLetter.UpdateScreen "On"
    End
End Sub
