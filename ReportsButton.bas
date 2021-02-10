Attribute VB_Name = "ReportsButton"
Option Explicit
Private acctList As Collection

Sub GenerateReports(Optional bothClicked As Boolean)
    LogData.TimeStart
    'Turn off screen updating to improve performance
    BothButton.UpdateScreen "Off"
    
    'Add vague error tracking
'    On Error GoTo MacroBroke
    
    'Set global variables
    If Not bothClicked Then
        'Just the report is being generated, not both report and LetterButton. Set the global variables
        '(TradeRecommendationsExport workbook and worksheet, household name, equity target, save location)
        BothButton.SetGlobals
    End If
    
    'Get a local version of AccountList
    Set acctList = BothButton.AccountList
    
    'Once there's a household, log macro use
    Dim macroName As String
    If BothButton.ReportBuilderSheet.Shapes("cbxTrades").OLEFormat.Object.Object.Value = True Then
        LogData.WriteLog "Report - Account", acctList(1).household, True
    End If
    If BothButton.ReportBuilderSheet.Shapes("cbxSubclass").OLEFormat.Object.Object.Value = True Then
        LogData.WriteLog "Report - Subclass", acctList(1).household, True
    End If
    
    'Generate the report(s) based on which checkboxes are checked
    If BothButton.ReportBuilderSheet.Shapes("cbxTrades").OLEFormat.Object.Object.Value = True Then
        TradeReport.BuildTradeReport
    End If
    If BothButton.ReportBuilderSheet.Shapes("cbxSubclass").OLEFormat.Object.Object.Value = True Then
        SubclassReport.BuildSubclassReport
    End If

    'Print the report(s) based on which checkboxes are checked
    If BothButton.ReportBuilderSheet.Shapes("cbxTrades").OLEFormat.Object.Object.Value = True Then
        PrintReport TradeReport.TrReport
    End If
    If BothButton.ReportBuilderSheet.Shapes("cbxSubclass").OLEFormat.Object.Object.Value = True Then
        PrintReport SubclassReport.SCReport
    End If
    
    'Save the exported csv as an excel workbook in the client's folder
    SaveReportWorkbook
    
    'Return everything back to normal
    Set TradeReport.TrReport = Nothing
    Set SubclassReport.SCReport = Nothing
    If Not bothClicked Then
        BothButton.ResetGlobals
    End If
    BothButton.UpdateScreen "On"
    
    'Log the time
    LogData.TimeEnd
    If Not bothClicked Then
        LogData.LogTime
    End If
    Exit Sub
MacroBroke:
    ErrorHandling.ErrorAndStop
End Sub

Sub FormatClientSheet(sht As Worksheet)
    With sht
        'Border at top of the page between content and header
        .Range("A1:F1").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A1:F1").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Cells.Font.Size = 11
        .Cells.Font.Name = "Arial"
        
        'Small space between header and content
        .Rows(1).RowHeight = 3
    End With
    
    'Format page layout
    AddHeader sht
    AddFooter sht
    FormatPrintArea sht
End Sub

Sub AddHeader(sht As Worksheet)
With sht
    'Get household name from BothButton
    Dim hhName As String
    If acctList.count > 0 Then
        hhName = acctList(1).household
    End If
    
    If hhName = "Schaefer, Russell & Patricia" Then
        'Client exception for the report
        .Range("A2").Value = "S, Russell & Patricia"
    Else
        'Put in the household name(s)
        .Range("A2").Value = hhName
    End If
    .Range("A2").Font.Bold = True
    
    If .Name = "Client Trades" Then
        .Range("A2:D2").Merge
        
        'Add the equity target
        .Range("E2").Value = "Equity Target"
        .Range("E2").Font.Underline = True
        .Range("E3").Value = BothButton.EqTarget
        .Range("E3").HorizontalAlignment = xlLeft
    ElseIf .Name = "Trades by Subclass" Then
        .Range("A2:C2").Merge
        
        'Add the equity target
        .Range("D2").Value = "Equity Target"
        .Range("D2").Font.Underline = True
        .Range("D3").Value = BothButton.EqTarget
        .Range("D3").HorizontalAlignment = xlLeft
    End If
        
    'Add the actual header
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
            ErrorHandling.ErrorAndContinue "FPIS logo not found. Header missing logo."
        Else
            .RightHeader = "&R&g"
            .RightHeaderPicture.fileName = imagePath
            .RightHeaderPicture.LockAspectRatio = msoTrue
            .RightHeaderPicture.Height = Application.InchesToPoints(0.7)
        End If
    End With
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
    BothButton.UpdateScreen "On"
    sht.DisplayPageBreaks = False
    LogData.TimeEnd
    sht.Columns("A:F").PrintOut Copies:=2, Collate:=True, Preview:=True
    LogData.TimeStart
End Sub

Sub SaveReportWorkbook()
    'Set the file name. Default file name is "[Month] [Year].xlsx"
    Dim saveFile As String
    saveFile = Trim(MonthName(Month(Date)) & " " & Year(Date))
    
    'Open save dialog in order to select save location
    LogData.TimeEnd
    Dim sfdSaveReport As FileDialog
    Set sfdSaveReport = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveReport.InitialFileName = BothButton.saveDir & saveFile
    sfdSaveReport.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveReport.SelectedItems.count > 0 Then
        savePath = sfdSaveReport.SelectedItems.Item(1)
        BothButton.ExportBook.SaveAs fileName:=savePath, FileFormat:=xlWorkbookDefault, ReadOnlyRecommended:=False
    End If
    LogData.TimeStart
End Sub

Function GetIndexOf(arr As Variant, str As String) As Integer
    'Return the index of an array's element
    Dim i As Integer
    GetIndexOf = -1
    i = 1
    Do While GetIndexOf = -1 And i <= UBound(arr)
        If arr(i) = str Then
            GetIndexOf = i
        End If
        i = i + 1
    Loop
    
    If GetIndexOf = -1 Then
        ErrorHandling.ErrorAndStop str & " not found on TradeRecommendationsExport sheet. Macro has been halted."
    End If
End Function

Function FindHeader(str As String) As Range
    'Call the actual FindHeader function
    BothButton.FindHeader (str)
End Function
