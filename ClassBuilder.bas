Attribute VB_Name = "ClassBuilder"
Option Explicit

Public Function NewAccount(Name As String, Number As String, acctType As String, Custodian As String) As clsAccount
    Dim tempAccount As clsAccount
    Set tempAccount = New clsAccount
    tempAccount.ClassConstructor Name, Number, acctType, Custodian
    Set NewAccount = tempAccount
End Function

Public Function NewCSV(Name As String) As clsCSV
    Dim tempCSV As clsCSV
    Set tempCSV = New clsCSV
    tempCSV.ClassConstructor windowNm:=Name
    Set NewCSV = tempCSV
End Function

Public Function NewHousehold() As clsHousehold
    Dim tempHousehold As clsHousehold
    Set tempHousehold = New clsHousehold
    tempHousehold.ClassConstructor
    Set NewHousehold = tempHousehold
End Function

Public Function NewLetter(household As clsHousehold) As clsLetter
    Dim tempLetter As clsLetter
    Set tempLetter = New clsLetter
    tempLetter.ClassConstructor household
    Set NewLetter = tempLetter
End Function

Public Function NewSubclass(TRXDescription As String, Description As String, accountList As Collection) As clsSubclass
    Dim tempSubclass As clsSubclass
    Set tempSubclass = New clsSubclass
    tempSubclass.ClassConstructor trx:=TRXDescription, desc:=Description, acctList:=accountList
    Set NewSubclass = tempSubclass
End Function

Public Function NewTradeRow(Symbol As String, Description As String, Subclass As String, Action As String, Amount As String, Percent As String) As clsTradeRow
    Dim tempRow As clsTradeRow
    Set tempRow = New clsTradeRow
    tempRow.ClassConstructor sym:=Symbol, desc:=Description, sc:=Subclass, act:=Action, am:=BlankToZero(Amount), pcnt:=BlankToZero(Percent)
    Set NewTradeRow = tempRow
End Function

Private Function BlankToZero(str As String) As Double
    If str = vbNullString Or Not IsNumeric(str) Then
        BlankToZero = 0
    Else
        BlankToZero = CDbl(str)
    End If
End Function

Public Function NewWindow(Name As String) As clsWindow
    Dim tempWindow As clsWindow
    Set tempWindow = New clsWindow
    tempWindow.ClassConstructor windowNm:=Name
    Set NewWindow = tempWindow
End Function
