VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTradeRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_symbol As String
Private m_description As String
Private m_subclass As String
Private m_action As String
Private m_amount As Double
Private m_percent As Double

Public Property Get Symbol() As String
    Symbol = m_symbol
End Property

Public Property Get Description() As String
    Description = m_description
End Property

Public Property Get Subclass() As String
    Subclass = m_subclass
End Property

Public Property Get Action() As String
    Action = m_action
End Property

Public Property Get Amount() As Double
    Amount = m_amount
End Property

Public Property Get Percent() As Double
    Percent = m_percent
End Property

Public Sub ClassConstructor(sym As String, desc As String, sc As String, act As String, am As String, pcnt As String)
    m_symbol = sym
    m_description = desc
    m_subclass = sc
    m_action = act
    m_amount = am
    m_percent = pcnt
End Sub

Public Sub AddToTrade(amountToAdd As Double)
    m_amount = m_amount + amountToAdd
End Sub

Public Sub ChangeAction(newAction As String)
    m_action = newAction
End Sub
