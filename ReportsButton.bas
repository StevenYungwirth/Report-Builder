Attribute VB_Name = "ReportsButton"
Option Explicit

Public Sub btnReport_Click()
    'Start the log
    LogData.TimeStart
    
    'Add vague error tracking
    If ThisWorkbook.Name <> "Test Report Builder.xlsm" Then
        On Error GoTo MacroBroke
    End If
    
    'Turn off screen updating to improve performance
    BothButton.UpdateScreen "Off"
    
    'Get the household
    Dim ClientHousehold As clsHousehold
    Set ClientHousehold = ClassBuilder.NewHousehold
    
    'Build the reports
    BuildReports ClientHousehold
    
    'Turn the screen updating back on
    BothButton.UpdateScreen "On"
    
    'Log the time
    LogData.TimeEnd
    LogData.LogTime
    
    Exit Sub
MacroBroke:
    If Not ClientHousehold Is Nothing Then
        ErrorHandling.ErrorAndStop hhName:=ClientHousehold.Name
    Else
        ErrorHandling.ErrorAndStop
    End If
End Sub

Public Sub BuildReports(household As clsHousehold)
    'Get the worksheet with the buttons
    Dim buttonWindow As New clsWindow
    Set buttonWindow = ClassBuilder.NewWindow("Report Builder")

    'Log macro use and generate the report(s) based on which checkboxes are checked
    Dim TrReport As Worksheet
    If buttonWindow.Book.Worksheets(1).Shapes("cbxTrades").OLEFormat.Object.Object.value = True Then
        'Build report that shows trades by account
        LogData.WriteLog "Report - Account", household.Name, True
        Set TrReport = TradeReport.BuildTradeReport(household)
    End If
    Dim scReport As Worksheet
    If buttonWindow.Book.Worksheets(1).Shapes("cbxSubclass").OLEFormat.Object.Object.value = True Then
        'Build report that shows trades by subclass
        LogData.WriteLog "Report - Subclass", household.Name, True
        Set scReport = SubclassReport.BuildSubclassReport(household)
    End If
    
    'Get the trade sheet in order to print the reports and save the workbook
    Dim tradeSheet As clsWindow
    Set tradeSheet = ClassBuilder.NewWindow("TradeRecommendationsExport")

    'Print the available reports
    Dim shtNum As Integer
    For shtNum = 1 To tradeSheet.Book.Worksheets.count
        Dim sht As Worksheet
        Set sht = tradeSheet.Book.Worksheets(shtNum)
        If InStr(tradeSheet.Book.Name, sht.Name) = 0 Then
            PrintReport sht
        End If
    Next shtNum
    
    'Save the exported csv as an excel workbook in the client's folder
    SaveReportWorkbook wkBook:=tradeSheet.Book, folderName:=household.ServerFolder
End Sub

Sub PrintReport(sht As Worksheet)
    'Check if the sheet exists; if not then it wasn't ran
    If Not sht Is Nothing Then
        'Enable screen updating and pause the timer
        BothButton.UpdateScreen "On"
        LogData.TimeEnd
        
        'Print the report
        sht.DisplayPageBreaks = False
        sht.Columns("A:F").PrintOut Preview:=True
        
        'Restart the timer and turn screen updating back off
        LogData.TimeStart
        BothButton.UpdateScreen "Off"
    End If
End Sub

Sub SaveReportWorkbook(wkBook As Workbook, folderName As String)
    'Set the file name. Default file name is "[Month] [Year].xlsx"
    Dim saveFile As String
    saveFile = Trim(MonthName(Month(Date)) & " " & Year(Date))
    
    'Enable screen updating and pause the timer
    BothButton.UpdateScreen "On"
    LogData.TimeEnd
        
    'Open save dialog in order to select save location
    Dim sfdSaveReport As FileDialog
    Set sfdSaveReport = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
    sfdSaveReport.InitialFileName = folderName & saveFile
    sfdSaveReport.Show
    
    'Save at selected location
    Dim saveSelected As Variant
    Dim savePath As String
    If sfdSaveReport.SelectedItems.count > 0 Then
        savePath = sfdSaveReport.SelectedItems.Item(1)
        wkBook.SaveAs fileName:=savePath, FileFormat:=xlWorkbookDefault, ReadOnlyRecommended:=False
    End If
    
    'Restart the timer and turn screen updating back off
    LogData.TimeStart
    BothButton.UpdateScreen "Off"
End Sub
