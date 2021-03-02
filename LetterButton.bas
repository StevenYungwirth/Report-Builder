Attribute VB_Name = "LetterButton"
Option Explicit

Public Sub btnLetter_Click()
    'Start the log
    LogData.TimeStart
    
    'Error handling
    If ThisWorkbook.Name <> "Test Report Builder.xlsm" Then
        On Error GoTo MacroBroke
    End If
    
    'Turn off screen updating
    BothButton.UpdateScreen "Off"
    
    'Get the household
    Dim ClientHousehold As clsHousehold
    Set ClientHousehold = ClassBuilder.NewHousehold
    
    'Build the letter
    BuildLetter household:=ClientHousehold
    
    'Turn the screen updating back on
    BothButton.UpdateScreen "On"
    
    'Log the time
    LogData.TimeEnd
    LogData.LogTime
    
    Exit Sub
MacroBroke:
    CloseWord
    If Not ClientHousehold Is Nothing Then
        ErrorHandling.ErrorAndStop hhName:=ClientHousehold.Name
    Else
        ErrorHandling.ErrorAndStop
    End If
End Sub

Public Sub BuildLetter(household As clsHousehold)
    'Log macro use
    LogData.WriteLog "Letter", household.Name, True
    
    'Get the letter
    Dim letter As clsLetter
    Set letter = ClassBuilder.NewLetter(household)
    
    'Add client names, advisor names, and rebalancing info to letter
    letter.Process
    
    'Print the letter
    letter.PrintLetter
    
    'Save the letter
    letter.SaveLetter
End Sub

Sub CloseWord()
    'Close Word document and application, and set them to nothing
    If Not WordDoc Is Nothing Then
        WordDoc.Close SaveChanges:=False
        Set WordDoc = Nothing
    End If
    If Not WordApp Is Nothing Then
        WordApp.Quit
        Set WordApp = Nothing
    End If
End Sub
