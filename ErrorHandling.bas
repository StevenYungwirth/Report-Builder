Attribute VB_Name = "ErrorHandling"
Option Explicit

Sub ErrorAndContinue(Optional message As String, Optional msgBoxTitle As String, Optional hhName As String)
    'Minor error. Show message and keep running macro
    ShowError message, msgBoxTitle
    LogData.LogError 1, hhName
End Sub
Sub ErrorAndStop(Optional message As String, Optional msgBoxTitle As String, Optional hhName As String)
    'Fatal error. Show message and stop the macro
    ShowError message, msgBoxTitle
    BothButton.UpdateScreen "On"
    LogData.LogError 2, hhName
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
    MsgBox message, Title:=msgBoxTitle
    
    'Restart the timer once the message box is closed
    LogData.TimeStart
End Sub

