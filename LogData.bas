Attribute VB_Name = "LogData"
Option Explicit
Private LogStarted As Boolean
Private TimerRunning As Boolean
Private StartingTime As Double
Private TotalElapsedTime As Double

Public Sub TimeStart()
    'Check if the timer's already running; don't change anything if it is
    If Not TimerRunning Then
        StartingTime = Now
        TimerRunning = True
        Debug.Print "Started at: " & StartingTime
    End If
End Sub

Public Sub TimeEnd()
    'Get the current time and compare it to when the timer started
    Dim ElapsedTime As Double
    ElapsedTime = Now - StartingTime
    TotalElapsedTime = TotalElapsedTime + ElapsedTime
    Debug.Print "Ended at: " & Now
    Debug.Print "Elapsed time: " & ElapsedTime
    Debug.Print "Total Elapsed time: " & Round(TotalElapsedTime * 60 * 60 * 24, 2)
    
    'Stop the timer
    TimerRunning = False
End Sub

Public Sub LogTime()
    'Get the total elapsed time in seconds
    Dim ElapsedTime As Double
    ElapsedTime = TotalElapsedTime * 60 * 60 * 24
    
    'Log the elapsed time
    If ElapsedTime > 60 Then
        'Log the elapsed time in minutes and seconds
        WriteLine chr(9) & "Total elapsed time: " & Round(ElapsedTime / 60, 0) & " minutes, " & Round(ElapsedTime - ElapsedTime / 60, 2) & " seconds"
    Else
        'Log the elapsed time in seconds
        WriteLine chr(9) & "Total elapsed time: " & Round(TotalElapsedTime * 60 * 60 * 24, 2) & " seconds"
    End If
    
    'Reset the timer
    TimerReset
End Sub

Private Sub TimerReset()
    LogStarted = False
    TimerRunning = False
    StartingTime = 0
    TotalElapsedTime = 0
End Sub

Public Sub WriteLog(macroName As String, household As String, runSuccess As Integer)
    'Get the header line
    Dim logStr As String
    logStr = HeaderLine
    
    'Add the macro used and the household to the log string
    If macroName = "Letter" Then
        'Add an extra tab
        logStr = chr(9) & "Macro: " & macroName & chr(9) & chr(9) & chr(9) & "Household: " & household
    Else
        logStr = chr(9) & "Macro: " & macroName & chr(9) & chr(9) & "Household: " & household
    End If
    
    'Write to the log
    WriteLine logStr
End Sub

Private Function HeaderLine() As String
    If Not LogStarted Then
        'The log hasn't started yet. Get the user name/computer as a header line
        HeaderLine = VBA.Environ("username") & chr(9) & VBA.Environ("computername") & " " & Now() & chr(13)
        
        'Start the log
        LogStarted = True
    End If
End Function

Private Sub WriteLine(strText As String, Optional logLocation As String)
    'Don't write to the log if this is the test builder
    If ThisWorkbook.Name <> "Test Report Builder.xlsm" Then
        'Set the file name for the log. File name is "yyyy-mm-dd.txt"
        Dim fileName As String
        fileName = Format(Date, "yyyy-mm-dd")
        
        'Open the log file
        If logLocation = vbNullString Then
            logLocation = "Z:\YungwirthSteve\Log\" & fileName & ".txt"
        End If
        Dim textFile As Integer
        textFile = FreeFile
        Open logLocation For Append As textFile
        
        'Add the text to the log
        Print #textFile, strText
        
        'Close the log file
        Close textFile
    End If
End Sub

Public Sub LogError(errorNum As Integer, hhName As String)
    'Get the user who had an error
    Dim errorLogString As String
    errorLogString = VBA.Environ("username") & chr(9) & VBA.Environ("computername") & " " & Now()
    errorLogString = errorLogString & chr(13) & chr(9)
    
    If errorNum = 1 Then
        WriteLine chr(9) & "Minor error occurred"
        errorLogString = errorLogString & "Minor error"
    ElseIf errorNum = 2 Then
        WriteLine chr(9) & "Fatal error occurred"
        errorLogString = errorLogString & "Fatal error"
    End If
    
    If hhName <> vbNullString Then
        errorLogString = errorLogString & " - " & hhName
    Else
        errorLogString = errorLogString & " - No household"
    End If
    
    WriteLine errorLogString, "Z:\YungwirthSteve\Log\ErrorLog.txt"
End Sub
