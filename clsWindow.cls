VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Book As Workbook
Private m_windowName As String

Public Property Get WindowName() As String
    WindowName = m_windowName
End Property

Public Sub ClassConstructor(windowNm As String)
    m_windowName = windowNm
    Set Book = GetBookByName(m_windowName)
End Sub

Private Function GetBookByName(fileName As String) As Workbook
    'Given a file name, loop through each open window to find one with that name
    Dim bookFound As Boolean
    Dim Window As Integer
    Window = 1
    Do While Not bookFound And Window < Windows.count
        'Compare the window and file names
        'Not checking for equality because there will often be multiple TradeRecommendationsExport files in the folder, so they would be indexed
        If InStr(UCase(Windows(Window).Caption), UCase(fileName)) > 0 Then
            'The window name contains the file name
            bookFound = True
        Else
            'The window name doesn't contain the file name, go to the next window
            Window = Window + 1
        End If
    Loop
    
    If Not bookFound Then
        'The file isn't open
        ErrorHandling.ErrorAndStop fileName & " not found"
    End If
    
    'Return the workbook
    Set GetBookByName = Workbooks(Windows(Window).Caption)
End Function
