VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_trxDescription As String
Private m_description As String
Private m_tradeList As New Collection

Public Property Get TRXDescription() As String
    TRXDescription = m_trxDescription
End Property
Public Property Get Description() As String
    Description = m_description
End Property

Public Property Get TradeList() As Collection
    Set TradeList = m_tradeList
End Property

Public Sub ClassConstructor(trx As String, desc As String, acctList As Collection)
    m_trxDescription = trx
    m_description = desc
    FillTradeList acctList
End Sub

Private Sub FillTradeList(accountList As Collection)
    'Loop through each account to get its trades
    Dim account As Variant
    For Each account In accountList
        'Loop through each trade and add it to the list if the subclass matches
        Dim trade As Variant
        For Each trade In account.TradeList
            'Add the security as a trade in its subclass
            If trade.Subclass = m_trxDescription And m_tradeList.count = 0 Then
                'The trade list is empty, add the fund
                m_tradeList.Add trade
            ElseIf trade.Subclass = m_trxDescription Then
                'See if the security is already in the list
                Dim isInList As Boolean
                Dim scFund As Variant
                For Each scFund In m_tradeList
                    If scFund.Symbol = trade.Symbol Then
                        'The security is already in the list, combine it with what's already there
                        scFund.AddToTrade trade.Amount
                        isInList = True
                    End If
                Next scFund
                
                If Not isInList Then
                    'The security wasn't already in the list, add it
                    m_tradeList.Add trade
                End If
                
                isInList = False
            End If
        Next trade
    Next account
End Sub
