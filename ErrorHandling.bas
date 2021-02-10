Attribute VB_Name = "ErrorHandling"
Option Explicit

Sub ErrorAndContinue(Optional message As String, Optional msgBoxTitle As String)
    'Minor error. Show message and keep running macro
    ShowError message, msgBoxTitle
    LogData.LogError 1
End Sub
Sub ErrorAndStop(Optional message As String, Optional msgBoxTitle As String)
    'Fatal error. Show message and stop the macro
    ShowError message, msgBoxTitle
    BothButton.UpdateScreen "On"
    LogData.LogError 2
    LogData.TimeEnd
    End
End Sub
Private Sub ShowError(message As String, msgBoxTitle As String)
    If message = vbNullString Then
        'ShowError was called without a message being provided (coding error)
        message = "An error occurred that wasn't accounted for."
        msgBoxTitle = "You found something new!"
    End If
    
    'Pause the timer
    LogData.TimeEnd
    
    'Show the error message
    If msgBoxTitle = vbNullString Then
        MsgBox message
    Else
        MsgBox message, Title:=msgBoxTitle
    End If
    
    'Restart the timer once the message box is closed
    LogData.TimeStart
End Sub
