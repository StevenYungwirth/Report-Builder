VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHousehold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const exportSheetName As String = "TradeRecommendationsExport"
Private m_clientFirst As String
Private m_clientLast As String
Private m_clientSuffix As String
Private m_SpouseFirst As String
Private m_SpouseLast As String
Private m_SpouseSuffix As String
Private m_tradeSheet As clsCSV
Private m_accounts As New Collection
Private m_serverFolder As String

Public Property Get Name() As String
    Name = m_tradeSheet.GetData("CRHouseholdDescription", 1)
End Property

Public Property Get ClientFirst() As String
    ClientFirst = m_clientFirst
End Property

Public Property Get ClientLast() As String
    ClientLast = m_clientLast
End Property

Public Property Get ClientSuffix() As String
    ClientSuffix = m_clientSuffix
End Property

Public Property Get SpouseFirst() As String
    SpouseFirst = m_SpouseFirst
End Property

Public Property Get SpouseLast() As String
    SpouseLast = m_SpouseLast
End Property

Public Property Get SpouseSuffix() As String
    SpouseSuffix = m_SpouseSuffix
End Property

Public Property Get tradeSheet() As clsCSV
    Set tradeSheet = m_tradeSheet
End Property

Public Property Get Allocation() As String
    Allocation = m_tradeSheet.GetData("AssetAllocationModel", 1)
End Property

Public Property Get eqTarget() As String
    Dim allo As String
    allo = Allocation
    
    'If there's a % in the string, remove it and everything to the right of it
    If InStr(allo, "%") <> 0 Then
        allo = Left(allo, InStr(allo, "%") - 1)
    End If
    
    'Remove characters until eqTarget is either blank or numeric
    Do While Not IsNumeric(allo) And allo <> ""
        If InStr(allo, " ") > 1 Then
            'There is a space
            If Not IsNumeric(Mid(allo, InStr(allo, " ") - 1, 1)) Then
                'The character before the space is not numeric, so remove it
                allo = Right(allo, Len(allo) - 1)
            Else
                'The character before the space is numeric
                If InStr(allo, "-") > 0 Or InStr(allo, "/") > 0 Then
                    If InStr(allo, "-") > InStr(allo, " ") Or InStr(allo, "/") > InStr(allo, " ") Then
                        'There's a special character after the space. Remove characters at the end until there's no space
                        Do While InStr(allo, " ") > 0
                            allo = Left(allo, Len(allo) - 1)
                        Loop
                    Else
                        'There's a special character before the space. Remove characters at the start until there's no space
                        Do While InStr(allo, " ") > 0
                            allo = Right(allo, Len(allo) - 1)
                        Loop
                    End If
                Else
                    'Default: The character before the space is numeric, but there's no special character. Remove characters at the end
                    allo = Left(allo, Len(allo) - 1)
                End If
            End If
        ElseIf InStr(allo, " ") = 1 Then
            'The space is the first character, so remove it
            allo = Right(allo, Len(allo) - 1)
        Else
            'No space, e.g. fpis80/80fpis
            If IsNumeric(Left(allo, 1)) Then
                'The first character is numeric, remove characters at the end
                allo = Left(allo, Len(allo) - 1)
            Else
                'The first character is not numeric, remove characters from the start
                allo = Right(allo, Len(allo) - 1)
            End If
        End If
    Loop
    
    If allo = "" Then
        'allo (presumably) contained only text
        eqTarget = ""
    Else
        eqTarget = Trim(allo) & "%"
    End If
End Property

Public Property Get Accounts() As Collection
    Set Accounts = m_accounts
End Property

Public Property Get ServerFolder() As String
    ServerFolder = m_serverFolder
End Property

Public Sub ClassConstructor()
    'Get the trades sheet
    Set m_tradeSheet = ClassBuilder.NewCSV(exportSheetName)
    
    'Load the accounts and sort the list by account name
    Set m_accounts = LoadAccounts
    Set m_accounts = SortAccountList
    
    'Parse the client names from the household name
    GetClientNames
    
    'Get the client's folder in the Z drive
    m_serverFolder = GetDir
End Sub

Private Function LoadAccounts() As Collection
    'Define a new collection
    Dim tempAccountList As Collection
    Set tempAccountList = New Collection
    
    'Get each account
    Set tempAccountList = FillAccountList
    
    'Put each account's trades into the account
    Dim acct As Variant
    For Each acct In tempAccountList
        acct.FillTradeList m_tradeSheet
    Next acct
    
    'Set the list to return
    Set LoadAccounts = tempAccountList
End Function

Private Function FillAccountList() As Collection
    'Instantiate a temporary collection
    Dim tempAccountList As Collection
    Set tempAccountList = New Collection

    'Loop through each trade row element to get each account
    Dim isInList As Boolean
    Dim ele As Integer
    For ele = 1 To UBound(m_tradeSheet.Data, 1) - 1
        isInList = False
        Dim acct As Variant
        For Each acct In tempAccountList
            If acct.Number = m_tradeSheet.GetData("AccountNumber", ele) Then
                'The account is in the list
                isInList = True
            End If
        Next acct
        
        'Add the account if it's not in the list
        If Not isInList Then
            'The account isn't in the list, create a new account
            Dim tempAccount As clsAccount
            Set tempAccount = ClassBuilder.NewAccount(Name:=m_tradeSheet.GetData("CRAccountMasterDescription", ele), _
                                Number:=m_tradeSheet.GetData("AccountNumber", ele), _
                                acctType:=m_tradeSheet.GetData("AccountType", ele), _
                                Custodian:=m_tradeSheet.GetData("Custodian", ele))
            
            'Add the new account to the list
            tempAccountList.Add tempAccount
        End If
    Next ele
    
    Set FillAccountList = tempAccountList
End Function

Private Function SortAccountList() As Collection
    'Instantiate a temporary list and add the first account
    Dim tempList As New Collection
    tempList.Add m_accounts(1)
    
    'If there's only one account, sorting is trivial
    If m_accounts.count >= 2 Then
        Dim i As Integer
        For i = 2 To m_accounts.count
            Dim isAdded As Boolean
            isAdded = False
            Dim j As Integer
            j = 1
            Do While Not isAdded And j < tempList.count
                If m_accounts(i).Name < tempList(j).Name Then
                    tempList.Add m_accounts(i), Before:=j
                    isAdded = True
                End If
                
                j = j + 1
            Loop
            
            'If the account wasn't put into the list, put it at the end
            If Not isAdded Then
                tempList.Add m_accounts(i)
            End If
        Next i
    End If
    
    'Return the sorted list
    Set SortAccountList = tempList
End Function

Private Sub GetClientNames()
    Dim householdName As String
    householdName = Name
    
    If InStr(householdName, ",") > 0 Then
        'Household name has a comma, find where it is
        Dim commaPos As Integer
        commaPos = InStr(householdName, ",")
        
        'Declare name arrays
        Dim clientNames(2, 1) As String
        
        'Take last name as everything before the comma
        clientNames(1, 0) = Trim(Left(householdName, commaPos - 1))
        clientNames(1, 0) = Replace(clientNames(1, 0), " ", "")
        clientNames(1, 0) = Replace(clientNames(1, 0), "-", "")
        
        'Get first name
        If InStr(householdName, "&") > 0 Then
            'Household name has spouse's name
            Dim ampPos As Integer
            ampPos = InStr(householdName, "&")

            'Everything between comma and ampersand is the first name
            clientNames(0, 0) = Trim(Mid(householdName, commaPos + 1, ampPos - commaPos - 1))
            
            'Everything after the ampersand is the spouse's first name
            clientNames(0, 1) = Trim(Right(householdName, Len(householdName) - ampPos - 1))
            
            If InStr(clientNames(0, 1), ",") > 0 Then
                'Spouse has a different last name
                Dim spouseCommaPos As Integer
                spouseCommaPos = InStr(clientNames(0, 1), ",")
                clientNames(1, 1) = Trim(Left(clientNames(0, 1), spouseCommaPos - 1))
                clientNames(1, 1) = Replace(clientNames(1, 1), " ", "")
                clientNames(1, 1) = Replace(clientNames(1, 1), "-", "")
                clientNames(0, 1) = Trim(Right(clientNames(0, 1), Len(clientNames(0, 1)) - spouseCommaPos))
            Else
                'Spouse has the same last name
                clientNames(1, 1) = clientNames(1, 0)
                clientNames(1, 1) = Replace(clientNames(1, 1), " ", "")
                clientNames(1, 1) = Replace(clientNames(1, 1), "-", "")
            End If
        Else
            'Household name is just one person, take everything after the comma as the first name
            clientNames(0, 0) = Trim(Right(householdName, Len(householdName) - commaPos))
            clientNames(0, 1) = ""
            clientNames(1, 1) = ""
        End If
        
        If InStr(clientNames(0, 0), " ") > 0 Then
            'Client's name has a suffix
            clientNames(2, 0) = Trim(Right(clientNames(0, 0), Len(clientNames(0, 0)) - InStr(clientNames(0, 0), " ")))
            clientNames(0, 0) = Trim(Left(clientNames(0, 0), InStr(clientNames(0, 0), " ")))
        End If
        If InStr(clientNames(0, 0), " ") > 0 Then
            'Spouse's name has a suffix
            clientNames(2, 1) = Trim(Right(clientNames(0, 1), Len(clientNames(0, 1)) - InStr(clientNames(0, 1), " ")))
            clientNames(0, 1) = Trim(Left(clientNames(0, 1), InStr(clientNames(0, 1), " ")))
        End If
        
        'Set the client/spouse names
        m_clientFirst = clientNames(0, 0)
        m_clientLast = clientNames(1, 0)
        m_clientSuffix = clientNames(2, 0)
        m_SpouseFirst = clientNames(0, 1)
        m_SpouseLast = clientNames(1, 1)
        m_SpouseSuffix = clientNames(2, 1)
    Else
        'TODO Household name has no comma
    End If
End Sub

Private Function GetDir() As String
    'Set the folder name
    Dim folderName As String
    folderName = ClientFolder
    
    'See if the folder is a valid location in the Z drive
    If FolderExists(folderName) Then
        'The folder is valid location. See if there's a letters folder within it
        If FolderExists(folderName & "\Letters\") Then
            'Full path is available, have the default save location be here
            GetDir = "Z:\" & folderName & "\Letters\"
        ElseIf FolderExists(folderName & "\Letter\") Then
            'Full path is available, have the default save location be here
            GetDir = "Z:\" & folderName & "\Letter\"
        Else
            'Have the default save location be the client's folder
            GetDir = "Z:\" & folderName & "\"
        End If
    Else
        'The client's folder can't be found, open the dialog at the Z drive
        GetDir = "Z:\"
    End If
End Function

Private Function ClientFolder() As String
    Dim folderName As String
    folderName = Name
    If FolderExists(folderName) Then
        'Good; folder name is household name
    Else
        'Get the name of the client's folder. Returns folderName="" if not found
        folderName = FindFolder()
    End If
    
    ClientFolder = folderName
End Function

Function FindFolder() As String
    'Both names are [first name][last name][suffix]
    Dim cName As String
    Dim sName As String
    cName = m_clientLast & m_clientFirst & m_clientSuffix
    sName = m_SpouseLast & m_SpouseFirst & m_SpouseSuffix
    
    'Find the first letter of the client's name
    Dim firstChr As String
    firstChr = UCase(Left(cName, 1))
    
    'Get the first folder that starts with the first letter
    Dim zFolders() As String
    ReDim zFolders(0 To 0)
    zFolders(0) = Dir("Z:\" & firstChr & "*", vbDirectory)
    
    'Get the next folder that starts with the first letter
    Dim nextDir As String
    nextDir = Dir()
    
    'Get all folder names that start with the first letter of the client's name
    Dim count As Integer
    count = 0
    Do While nextDir <> ""
        count = count + 1
        ReDim Preserve zFolders(0 To count)
        zFolders(count) = nextDir
        nextDir = Dir()
    Loop
    
    'Narrow down the array by starting with the first two letters of the client's last name
    'Reduce the options until the folder is found or the end of the name is reached
    Dim chr As Integer
    chr = 2
    
    Dim folderName As String
    Dim exitLoop As Boolean
    exitLoop = False
    Do While Not exitLoop And chr <= Len(cName)
        zFolders = ReduceArr(cName, zFolders, chr)
        If UBound(zFolders) = 0 Then
            'The array only has one entry left, no need to go further
            folderName = zFolders(0)
            exitLoop = True
        End If
        chr = chr + 1
    Loop
    
    'Array could have multiple values in it after this
    If UBound(zFolders) > 0 Then
        If m_SpouseFirst <> "" Then
            'Folder isn't under client's name, try the spouse's name
            Dim emptyArr(2) As String
            folderName = FindFolder()
        Else
            'Spouse's name was tried already
        End If
    End If
    
    If Not FolderExists(folderName) Then
        folderName = NameCombinations()
    End If
    FindFolder = folderName
End Function

Private Function ReduceArr(Name As String, arr() As String, chr As Integer) As String()
    Dim xChr As String
    xChr = UCase(Left(Name, chr))
    
    Dim outArr() As String
    Dim xCount As Integer
    Dim ele As Variant
    Dim eleName As String
    ReDim outArr(0) As String
    For Each ele In arr
        If Left(UCase(ele), chr) = xChr Then
            ReDim Preserve outArr(0 To xCount) As String
            outArr(xCount) = ele
            xCount = xCount + 1
        End If
    Next ele

    ReduceArr = outArr
End Function

Private Function FolderExists(dirStr As String) As Boolean
    'Return whether or not the folder name is found in the Z drive
    Dim saveDir As String
    saveDir = "Z:\" & dirStr & "\"
    If Dir(saveDir, vbDirectory) = "" Or dirStr = "" Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

Private Function NameCombinations() As String
    'Declare arrays for every possible character or word in every position
    Dim posZero() As String
    Dim posOne() As String
    Dim posTwo() As String
    Dim posThree() As String
    Dim posFour() As String
    Dim posFive() As String
    Dim posSix() As String
    Dim posSeven() As String
    
    'Not as many options for the folder name if there's only one person
    If m_SpouseFirst = "" Then
        ReDim posFour(0) As String
        ReDim posFive(0) As String
        ReDim posSix(0) As String
        ReDim posSeven(0) As String
        'last
        posZero = Split(m_clientLast, ";")
        'blank, space, comma, first
        posOne = Split("; ;,;" & m_clientFirst, ";")
        'blank, space, first
        posTwo = Split("; ;" & m_clientFirst, ";")
        'blank, first
        posThree = Split(";" & m_clientFirst, ";")
        'blank
        posFour(0) = ""
        'blank
        posFive(0) = ""
        'blank
        posSix(0) = ""
        'blank
        posSeven(0) = ""
    Else
        'last, spouse last
        posZero = Split(m_clientLast & ";" & m_SpouseLast, ";")
        'blank, space, comma, first, spouse first, spouse last
        posOne = Split("; ;,;" & m_clientFirst & ";" & m_SpouseFirst & ";" & m_SpouseLast, ";")
        'blank, space, &, and, first, spouse first
        posTwo = Split("; ;&;and;" & m_clientFirst & ";" & m_SpouseFirst, ";")
        'blank, space, &, and, first, spouse first
        posThree = Split("; ;&;and;" & m_clientFirst & ";" & m_SpouseFirst, ";")
        'blank, space, &, and, first, spouse first
        posFour = Split("; ;&;and;" & m_clientFirst & ";" & m_SpouseFirst, ";")
        'blank, space, &, and, first, spouse first
        posFive = Split("; ;&;and;" & m_clientFirst & ";" & m_SpouseFirst, ";")
        'blank, space, first, spouse first
        posSix = Split("; ;" & m_clientFirst & ";" & m_SpouseFirst, ";")
        'blank, first, spouse first
        posSeven = Split(";" & m_clientFirst & ";" & m_SpouseFirst, ";")
    End If
    
    'Run through every combination of possible folder names, but leave the loops once the folder name is found
    Dim folderName As String
    Dim found As Boolean
    Dim zeroth, first, second, third, fourth, fifth, sixth, seventh As Integer
    found = False
    
    seventh = 0
    Do While Not found And seventh <= UBound(posSeven)
        sixth = 0
        Do While Not found And sixth <= UBound(posSix)
            fifth = 0
            Do While Not found And fifth <= UBound(posFive)
                fourth = 0
                Do While Not found And fourth <= UBound(posFour)
                    third = 0
                    Do While Not found And third <= UBound(posThree)
                        second = 0
                        Do While Not found And second <= UBound(posTwo)
                            first = 0
                            Do While Not found And first <= UBound(posOne)
                                zeroth = 0
                                Do While Not found And zeroth <= UBound(posZero)
                                    folderName = posZero(zeroth) & posOne(first) & posTwo(second) & posThree(third) & posFour(fourth) & posFive(fifth) & posSix(sixth) & posSeven(seventh)
                                    found = FolderExists(folderName)
                                    zeroth = zeroth + 1
                                Loop
                                first = first + 1
                            Loop
                            second = second + 1
                        Loop
                        third = third + 1
                    Loop
                    fourth = fourth + 1
                Loop
                fifth = fifth + 1
            Loop
            sixth = sixth + 1
        Loop
        seventh = seventh + 1
    Loop
    
    If Not found Then
        NameCombinations = ""
    Else
        NameCombinations = folderName
    End If
End Function
