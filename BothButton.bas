Attribute VB_Name = "BothButton"
Option Explicit
Public Sub BuildBoth()
    'Start the timer
    LogData.TimeStart
    
    'Turn off screen updating for better performance
    UpdateScreen "Off"
    
    'TODO Make error tracking better
    If ThisWorkbook.Name <> "Test Report Builder.xlsm" Then
        On Error GoTo MacroBroke
    End If
    
    'Get the household
    Dim ClientHousehold As clsHousehold
    Set ClientHousehold = ClassBuilder.NewHousehold
    
    'Log macro use
    LogData.WriteLog "Both Button", ClientHousehold.Name, True
    
    'Build both letter and report and print off each
    ReportsButton.BuildReports household:=ClientHousehold
    LetterButton.BuildLetter household:=ClientHousehold
    
    'Turn the screen updating back on
    UpdateScreen "On"
    
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

Public Sub UpdateScreen(OnOrOff As String)
    Dim reset As Long
    If OnOrOff = "Off" Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
    ElseIf OnOrOff = "On" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        reset = ActiveSheet.UsedRange.Rows.count
    End If
End Sub
