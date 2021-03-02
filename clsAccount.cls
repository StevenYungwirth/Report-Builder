VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_number As String
Private m_acctType As String
Private m_custodian As String
Private m_tradeList As New Collection

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Get Number() As String
    Number = m_number
End Property

Public Property Get acctType() As String
    acctType = m_acctType
End Property

Public Property Get Custodian() As String
    Custodian = m_custodian
End Property

Public Property Get TradeList() As Collection
    Set TradeList = m_tradeList
End Property

Public Sub ClassConstructor(nm As String, nmbr As String, aType As String, cust As String)
    m_name = nm
    m_number = nmbr
    m_acctType = aType
    m_custodian = cust
End Sub

Public Sub FillTradeList(tradeSheet As clsCSV)
    'Loop through each trade row element to add the account's trades
    Dim ele As Integer
    For ele = 1 To UBound(tradeSheet.Data, 1) - 1
        If tradeSheet.GetData("AccountNumber", ele) = m_number Then
            'The row corresponds to this account. Set the trade row as a new trade row object
            Dim tempRow As clsTradeRow
            Set tempRow = ClassBuilder.NewTradeRow(Symbol:=UCase(tradeSheet.GetData("Symbol", ele)), _
                                                    Description:=tradeSheet.GetData("Description", ele), _
                                                    Subclass:=tradeSheet.GetData("SubClass", ele), _
                                                    Action:=tradeSheet.GetData("Action", ele), _
                                                    Amount:=tradeSheet.GetData("Trade", ele), _
                                                    Percent:=tradeSheet.GetData("PCNTSOLD", ele))
            
            'Add the trade to the account
            m_tradeList.Add tempRow
        End If
    Next ele
    
    'Replace cash with the appropriate money market
    CashToMM
    
    'Combine sells from different lots and delete extra rows
    CombineSameSymbols
    
    'Sort the account's transactions
    SortTrans
End Sub

Private Sub CashToMM()
    'If "CASH" in the list of transactions, combine it with the money market
    'Find if "CASH" is a symbol in the transaction list
    Dim cashIndex As Integer
    cashIndex = CashFoundAt
    If cashIndex <> -1 Then
        '"CASH" was found. Get the account's money market symbol
        Dim mmSymbol As String
        mmSymbol = GetMMSymbol
        If mmSymbol = vbNullString Then
            'There's only a cash transaction; no other money market. Change "CASH" to be the default
            TradeList(cashIndex).Symbol = "MMDA12"
            TradeList(cashIndex).Description = "TD BANK FDIC Insured Money Market"
        Else
            'Combine the cash transaction with the money market
            Dim symb As Variant
            For Each symb In TradeList
                If symb.Symbol = mmSymbol Then
                    'Add the cash trade to the money market trade
                    symb.trade = symb.trade + TradeList(cashIndex).trade
                End If
            Next symb
            
            'Remove the cash transaction from the trade list
            TradeList.Remove cashIndex
        End If
    End If
End Sub

Private Function CashFoundAt() As Integer
    CashFoundAt = -1
    Dim trans As Integer
    For trans = 1 To TradeList.count
        If TradeList(trans).Symbol = "CASH" Then
            'Return the row the cash was found at
            CashFoundAt = trans
        End If
    Next trans
End Function

Private Function GetMMSymbol() As String
    'Return the symbol for the primary money market fund in the account
    Dim trade As Variant
    For Each trade In TradeList
        If trade.Subclass = "MMM" And trade.Symbol <> "CASH" Then
            GetMMSymbol = trade.Symbol
        End If
    Next trade
End Function

Private Sub CombineSameSymbols()
    'A security can have sales across many lots. Combine them all together
    'For each transaction, see if there is another transaction of the same symbol
    Dim trans As Integer
    trans = 1
    Do While trans <= TradeList.count
        Dim initialTrade As clsTradeRow
        Set initialTrade = TradeList(trans)
        'Go through transactions backwards so they can be removed if they have the same symbol
        Dim i As Integer
        For i = TradeList.count To trans + 1 Step -1
            Dim dupeTrade As clsTradeRow
            Set dupeTrade = TradeList(i)
            If initialTrade.Symbol = dupeTrade.Symbol Then
                'The transaction has the same symbol. Add the trade amount to the first one
                initialTrade.AddToTrade dupeTrade.Amount
                
                'Remove the extra transaction
                m_tradeList.Remove i
            End If
        Next i
        
        'If we're selling 100%, change the action to "SELL ALL"
        If initialTrade.Percent = 1 Then
            initialTrade.ChangeAction "SELL ALL"
        End If
            
        trans = trans + 1
    Loop
End Sub

Private Sub SortTrans()
    'Create a temporary list to hold the sorted transactions
    Dim tempTrans As Collection
    Set tempTrans = New Collection
    
    'Fill the temporary list with the transactions, alphabetically by action and then by symbol
    Dim trans As Integer
    trans = m_tradeList.count
    Do While trans > 0
        Dim firstTrans As Integer
        firstTrans = trans
        
        'See if there's another transaction that should go before firstTrans
        Dim i As Integer
        For i = 1 To TradeList.count - 1
            If m_tradeList(i).Action > m_tradeList(firstTrans).Action Then
                firstTrans = i
            ElseIf m_tradeList(i).Action = m_tradeList(firstTrans).Action Then
                'If the actions are the same, take the transaction with the first alphabetical symbol
                If m_tradeList(i).Symbol < m_tradeList(firstTrans).Symbol Then
                    firstTrans = i
                End If
            End If
        Next i
        
        'Put the transaction into the temporary list and remove it from the account's transactions
        tempTrans.Add m_tradeList(firstTrans)
        m_tradeList.Remove firstTrans
        
        trans = trans - 1
    Loop
    
    'Put the sorted list back into the current account
    Set m_tradeList = tempTrans
End Sub
