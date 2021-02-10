Attribute VB_Name = "LogData"
Option Explicit
Public LogStarted As Boolean
Private TimerRunning As Boolean
Private StartingTime As Double
Private TotalElapsedTime As Double

Sub WriteLog(macroName As String, household As String, runSuccess As Integer)
    If Not LogData.LogStarted Then
        'Log the user name/computer as a header line
        StartLog
    End If
    
    'Log the macro used and the household
    Dim logStr As String
    If macroName = "Letter" Then
        'Add an extra tab
        logStr = chr(9) & "Macro: " & macroName & chr(9) & chr(9) & chr(9) & "Household: " & household
    Else
        logStr = chr(9) & "Macro: " & macroName & chr(9) & chr(9) & "Household: " & household
    End If
    WriteLine logStr
End Sub
Private Sub StartLog()
    'Log the user name/computer as a header line
    Dim logStr As String
    logStr = VBA.Environ("username") & chr(9) & VBA.Environ("computername") & " " & Now()
    WriteLine logStr
    
    'Set LogStarted to be true
    LogStarted = True
End Sub
Private Sub WriteLine(strText As String)
    'Get date for log file name
    Dim todayDate As String
    todayDate = Format(Date, "yyyy-mm-dd")
    
    'Open the log file
    Dim LogFileLocation As String
    LogFileLocation = "Z:\YungwirthSteve\Log\" & todayDate & ".txt"
    Dim textFile As Integer
    textFile = FreeFile
    Open LogFileLocation For Append As textFile
    
    'Add the text to the log
    Print #textFile, strText
    
    'Close the log file
    Close textFile
End Sub
Sub LogTime()
    Dim ElapsedTime As Double
    ElapsedTime = TotalElapsedTime * 60 * 60 * 24
    If ElapsedTime > 60 Then
        WriteLine chr(9) & "Total elapsed time: " & Round(ElapsedTime / 60, 0) & " minutes, " & Round(ElapsedTime - ElapsedTime / 60, 2) & " seconds"
    Else
        WriteLine chr(9) & "Total elapsed time: " & Round(TotalElapsedTime * 60 * 60 * 24, 2) & " seconds"
    End If
    TotalElapsedTime = 0
End Sub
Sub LogError(errorNum As Integer)
    If errorNum = 1 Then
        WriteLine chr(9) & "Minor error occurred"
    ElseIf errorNum = 2 Then
        WriteLine chr(9) & "Fatal error occurred"
    End If
End Sub
Sub TimeStart()
    If Not TimerRunning Then
        StartingTime = Now
        TimerRunning = True
        Debug.Print "Started at: " & StartingTime
    Else
        'The timer is already running; do nothing
    End If
End Sub
Sub TimeEnd()
    Dim EndingTime As Double
    EndingTime = Now
    Dim ElapsedTime As Double
    ElapsedTime = EndingTime - StartingTime
    TotalElapsedTime = TotalElapsedTime + ElapsedTime
    TimerRunning = False
    Debug.Print "Ended at: " & EndingTime
    Debug.Print "Elapsed time: " & ElapsedTime
    Debug.Print "Total Elapsed time: " & Round(TotalElapsedTime * 60 * 60 * 24, 2)
End Sub
