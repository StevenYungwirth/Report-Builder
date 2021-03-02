VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_window As clsWindow
Private m_data() As Variant
Private m_headers() As String

Public Property Get Window() As clsWindow
    Set Window = m_window
End Property

Public Property Get Data() As Variant
    Data = m_data
End Property

Public Sub ClassConstructor(windowNm As String)
    Set m_window = ClassBuilder.NewWindow(windowNm)
    m_data = m_window.Book.Worksheets(1).UsedRange.Value2
    m_headers = GetCSVHeaders
End Sub

Private Function GetCSVHeaders() As String()
    'Set the array to be the values of the first row
    Dim headers() As String
    ReDim headers(1 To UBound(m_data, 2)) As String
    Dim col As Integer
    For col = 1 To UBound(m_data, 2)
        headers(col) = Trim(m_data(1, col))
    Next col
    
    'Check to make sure these column headers can be found. If one's not available, throw an error
    Dim neededHeaders() As Variant
    neededHeaders = Array("AccountNumber", "CRAccountMasterDescription", "Custodian", "Symbol", "OriginalTradeDate", "CostBasis", "Trade", _
                        "AccountType", "Action", "Description", "PCNTSOLD")
    Dim i As Integer
    For i = 0 To UBound(neededHeaders)
        'Check if the header is in the list of headers
        If Not DoesHeaderExist(headers, neededHeaders(i)) Then
            'The needed column wasn't found
            If DoesHeaderExist(headers, "CRHouseholdDescription") Then
                'Show the household name with the error
                ErrorHandling.ErrorAndStop neededHeaders(i) & " not found on TradeRecommendationsExport sheet. Macro has been halted.", hhName:=GetData("CRHouseholdDescription", 1)
            Else
                ErrorHandling.ErrorAndStop neededHeaders(i) & " not found on TradeRecommendationsExport sheet. Macro has been halted."
            End If
        End If
    Next i
    
    'Return the array
    GetCSVHeaders = headers
End Function

Private Function DoesHeaderExist(headerArr As Variant, Header As Variant) As Boolean
    'If the header doesn't exist, GetIndexOf will return -1
    If GetIndexOf(headerArr, Header) = -1 Then
        DoesHeaderExist = False
    Else
        DoesHeaderExist = True
    End If
End Function

Private Function GetIndexOf(arr As Variant, str As Variant) As Integer
    'Return the index of an array's element, or -1 in case the element can't be found
    Dim i As Integer
    GetIndexOf = -1
    i = 1
    Do While GetIndexOf = -1 And i <= UBound(arr)
        If arr(i) = str Then
            GetIndexOf = i
        End If
        i = i + 1
    Loop
End Function

Public Function GetData(headerName As String, datapoint As Integer) As String
    'Return the nth data point of a given header
    Dim dataFound As Boolean
    Dim col As Integer
    col = 1
    Do While Not dataFound And col < UBound(m_data, 2)
        'Look for the inputted header
        If Trim(m_data(1, col)) = headerName Then
            'Get the data. Datapoint is +1 since the first row is the header
            GetData = m_data(datapoint + 1, col)
            dataFound = True
        End If
        
        'Look for the header in the next column
        col = col + 1
    Loop
End Function

