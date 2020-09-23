Attribute VB_Name = "modUserDB"
' Replace these with your own variables
Const mvarUserDB$ = "C:\test.udb"
Const mvarPassword$ = "GROK"
' --------------------------------------


Public Enum AccessFunction
    afAddFlag = 1
    afRemFlag = 2
    afEditFlag = 3
    afEnumerate = 4
    afGetDescription = 5
End Enum
Public Enum CPFunction
    cpAddCheckpoint = 1
    cpRemCheckpoint = 2
    cpEditCheckpoint = 3
    cpEnumerate = 4
    cpGetFlags = 5
End Enum
Public Enum ULFunction
    ulAddUser = 1
    ulRemUser = 2
    ulEditUser = 3
    ulEnumerate = 4
    ulGetFlags = 5
    ulSetPass = 6
    ulGetRealName = 7
    ulGetPass = 8
    ulUserExists = 9
End Enum
Public Enum PType
    ptRequired = 1
    ptProhibited = 2
End Enum
Function ParseCPFlags(ParseType As PType, PString As String)
If ParseType = ptRequired Then
    F1 = ParseString(PString, 0, "-")
    If F1 = "NOT FOUND" Then ParseCPFlags = "" Else ParseCPFlags = F1
ElseIf ParseType = ptProhibited Then
    F2 = ParseString(PString, 1, "-")
    If F2 = "NOT FOUND" Then ParseCPFlags = "" Else ParseCPFlags = F2
End If
End Function
Public Sub MatchUser(DestVar As Collection, Optional UserID As String, Optional RealName As String, Optional UserAccess As String)
    On Error GoTo err_Init
    SQLz = ""
    If UserID <> "" Then SQLz = SQLz & "AND UserID LIKE '" & UserID & "' "
    If RealName <> "" Then SQLz = SQLz & "AND Realname LIKE '" & RealName & "' "
    If UserAccess <> "" Then SQLz = SQLz & "AND Access LIKE '" & UserAccess & "' "
    
    Dim UserDataBase As Database
    Dim UserDBRecordset As Recordset
    strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
    Set UserDataBase = OpenDatabase("", False, False, strn)
    Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE Password <> '' " & SQLz, dbOpenDynaset)
    If UserDBRecordset.RecordCount = 0 Then Exit Sub
    UserDBRecordset.MoveFirst
    For pnCnt = DestVar.Count To 1 Step -1
        DestVar.Remove pnCnt
    Next
    While Not UserDBRecordset.EOF
        DestVar.Add CStr(UserDBRecordset!UserID)
        UserDBRecordset.MoveNext
    Wend
End Sub
    Exit Sub

err_Init:
    MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
    
Public Function AddUser(UserID As String, RealName As String, UserAccess As String, Password As String) As Boolean
    On Error GoTo err_Init
    Call CreateDB(mvarUserDB, mvarPassword)
    strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
    Dim UserDataBase As Database
    Dim UserDBRecordset As Recordset
    Set UserDataBase = OpenDatabase("", False, False, strn)
    Set UserDBRecordset = UserDataBase.OpenRecordset("User", dbOpenDynaset)
    UserDBRecordset.AddNew
        UserDBRecordset!UserID = UserID
        UserDBRecordset!RealName = RealName
        UserDBRecordset!access = UserAccess
        UserDBRecordset!Password = Password
    UserDBRecordset.Update
    AddUser = True
    Exit Function
err_Init:
    MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
    AddUser = False
End Function
Public Function VerifyPassword(UserID As String, Password As String) As Boolean
    On Error GoTo err_Init
    strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
    Dim UserDataBase As Database
    Dim UserDBRecordset As Recordset
    Set UserDataBase = OpenDatabase("", False, False, strn)
    Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User Where UserID = '" & UserID & "'", dbOpenDynaset)
    If UserDBRecordset.RecordCount = 0 Then VerifyPassword = False
    If UserDBRecordset!Password = Password Then VerifyPassword = True Else VerifyPassword = False

    Exit Function
err_Init:
    MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
    VerifyPassword = False
End Function

Public Function AccessFlag(aCommand As AccessFunction, Optional fFlag As String, Optional fDescription As String, Optional NewFlag As String, Optional NewDescription As String, Optional enumCollection As Collection)
On Error GoTo err_Init
strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
Dim UserDataBase As Database
Dim UserDBRecordset As Recordset
Set UserDataBase = OpenDatabase("", False, False, strn)
Set UserDBRecordset = UserDataBase.OpenRecordset("Access", dbOpenDynaset)

                      
If aCommand = afAddFlag Then
        If fFlag = "" Then Call Err.Raise(1, , "Cannot process function without a Flag"): Exit Function
        If fDescription = "" Then Call Err.Raise(2, , "Cannot process function without a Description"): Exit Function
        Call AddFlag(fFlag, fDescription)
        AccessFlag = True
ElseIf aCommand = afEnumerate Then
        For pnCnt = enumCollection.Count To 1 Step -1
            enumCollection.Remove pnCnt
        Next
        While Not UserDBRecordset.EOF
            enumCollection.Add CStr(UserDBRecordset!flag)
            UserDBRecordset.MoveNext
        Wend
ElseIf aCommand = afGetDescription Then
        If fFlag = "" Then Call Err.Raise(1, , "Cannot process function without a Flag"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM Access WHERE Flag = '" & fFlag & "'", dbOpenDynaset)
        AccessFlag = UserDBRecordset!Description
ElseIf aCommand = afRemFlag Then
        If fFlag = "" Then Call Err.Raise(1, , "Cannot process function without a Flag"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM Access WHERE Flag = '" & fFlag & "'", dbOpenDynaset)
        UserDBRecordset.Delete
        AccessFlag = True
ElseIf aCommand = afEditFlag Then
        If fFlag = "" Then Call Err.Raise(1, , "Cannot process function without a Flag"): Exit Function
        If NewFlag = "" Then Call Err.Raise(1, , "Cannot process function without a NewFlag"): Exit Function
        If NewDescription = "" Then Call Err.Raise(1, , "Cannot process function without a NewDescription"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM Access WHERE Flag = '" & fFlag & "'", dbOpenDynaset)
        UserDBRecordset.Edit
        If NewFlag <> "" Then UserDBRecordset!flag = NewFlag Else UserDBRecordset!flag = UserDBRecordset!flag
        If NewDescription <> "" Then UserDBRecordset!Description = NewDescription Else UserDBRecordset!Description = UserDBRecordset!Description
        UserDBRecordset.Update
        AccessFlag = True
End If
    Exit Function

err_Init:
        MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
        AccessFlag = False
End Function
Private Sub AddCP(CP, flags)
    On Error GoTo err_Init
    Call CreateDB(mvarUserDB, mvarPassword)
    strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
    Dim UserDataBase As Database
    Dim UserDBRecordset As Recordset
    Set UserDataBase = OpenDatabase("", False, False, strn)
    Set UserDBRecordset = UserDataBase.OpenRecordset("Checkpoint", dbOpenDynaset)
    UserDBRecordset.AddNew
        UserDBRecordset!CheckPoint = CP
        UserDBRecordset!flag = flags
    UserDBRecordset.Update
    Exit Sub
err_Init:
     MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
End Sub

Private Sub AddFlag(flag, Desc)
    On Error GoTo err_Init
    Call CreateDB(mvarUserDB, mvarPassword)
    strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
    Dim UserDataBase As Database
    Dim UserDBRecordset As Recordset
    Set UserDataBase = OpenDatabase("", False, False, strn)
    Set UserDBRecordset = UserDataBase.OpenRecordset("Access", dbOpenDynaset)
    UserDBRecordset.AddNew
        UserDBRecordset!flag = flag
        UserDBRecordset!Description = Desc
    UserDBRecordset.Update
    Exit Sub
err_Init:
     MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
End Sub
Public Function CheckPoint(aCommand As CPFunction, Optional CPoint As String, Optional CFlags As String, Optional NewCPoint As String, Optional NewFlags As String, Optional enumCollection As Collection)
On Error GoTo err_Init
strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
Dim UserDataBase As Database
Dim UserDBRecordset As Recordset
Set UserDataBase = OpenDatabase("", False, False, strn)
Set UserDBRecordset = UserDataBase.OpenRecordset("Checkpoint", dbOpenDynaset)

              
If aCommand = cpAddCheckpoint Then
        If CPoint = "" Then Call Err.Raise(1, , "Cannot process function without a CPoint"): Exit Function
        If CFlags = "" Then Call Err.Raise(2, , "Cannot process function without a Flag"): Exit Function
        Call AddCP(CPoint, CFlags)
        CheckPoint = True
ElseIf aCommand = cpEnumerate Then
        For pnCnt = enumCollection.Count To 1 Step -1
            enumCollection.Remove pnCnt
        Next
        While Not UserDBRecordset.EOF
            enumCollection.Add CStr(UserDBRecordset!CheckPoint)
            UserDBRecordset.MoveNext
        Wend
ElseIf aCommand = cpGetFlags Then
        If CPoint = "" Then Call Err.Raise(1, , "Cannot process function without a CPoint"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM Checkpoint WHERE Checkpoint = '" & CPoint & "'", dbOpenDynaset)
        'If UserDBRecordset.RecordCount = 0 Then CheckPoint = ""
        CheckPoint = UserDBRecordset!flag
ElseIf aCommand = cpRemCheckpoint Then
        If CPoint = "" Then Call Err.Raise(1, , "Cannot process function without a CPoint"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM Checkpoint WHERE Checkpoint = '" & CPoint & "'", dbOpenDynaset)
        UserDBRecordset.Delete
        CheckPoint = True
ElseIf aCommand = cpEditCheckpoint Then
        If CPoint = "" Then Call Err.Raise(1, , "Cannot process function without a CPoint"): Exit Function
        If NewCPoint = "" Then Call Err.Raise(1, , "Cannot process function without a NewCP"): Exit Function
        If NewFlags = "" Then Call Err.Raise(1, , "Cannot process function without a NewFlag"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM Checkpoint WHERE Checkpoint = '" & CPoint & "'", dbOpenDynaset)
        UserDBRecordset.Edit
        UserDBRecordset!CheckPoint = NewCPoint
        UserDBRecordset!flag = NewFlags
        UserDBRecordset.Update
        CheckPoint = True
End If
    Exit Function

err_Init:
        MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
        CheckPoint = False
End Function
' Thanks to Ian I. for this :)
Function ParseString(ByVal strString As String, ByVal intNthOccurance As Integer, ByVal strSeperator As String) As String
    Dim intIndex As Integer
    Dim intStartOfString As Integer
    Dim intEndOfString As Integer
    Dim boolNotFound As Integer
    
    'check for intNthOccurance = 0--i.e. fir
    '     st one


    If (intNthOccurance = 0) Then


        If (InStr(strString, strSeperator) > 0) Then
                ParseString = Left(strString, InStr(strString, strSeperator) - 1)
        Else
                ParseString = strString
        End If
    Else
        'not the first one
        'init start of string on first comma
        intStartOfString = InStr(strString, strSeperator)
        
        'place start of string after intNthOccur
        '     ance-th comma (-1 since
        'already did one
        boolNotFound = 0


        For intIndex = 1 To intNthOccurance - 1
            'get next comma
            intStartOfString = InStr(intStartOfString + 1, strString, strSeperator)
            'check for not found


            If (intStartOfString = 0) Then
                boolNotFound = 1
            End If
        Next intIndex
        
        'put start of string past 1st comma
        intStartOfString = intStartOfString + 1
        
        'check for ending in a comma


        If (intStartOfString > Len(strString)) Then
            boolNotFound = 1
        End If
        


        If (boolNotFound = 1) Then
            ParseString = "NOT FOUND"
        Else
            intEndOfString = InStr(intStartOfString, strString, strSeperator)
            
            ' check for no second comma (i.e. end of
            '     string)


            If (intEndOfString = 0) Then
                intEndOfString = Len(strString) + 1
            Else
                intEndOfString = intEndOfString - 1
            End If
            ParseString = Mid$(strString, intStartOfString, intEndOfString - intStartOfString + 1)
        End If
    End If
End Function

Public Function UserList(aCommand As ULFunction, Optional UserID As String, Optional RealName As String, Optional flags As String, Optional Password As String, Optional NewID As String, Optional NewName As String, Optional NewFlags As String, Optional NewPass As String, Optional enumCollection As Collection)
On Error GoTo err_Init
strn = ";DATABASE=" & mvarUserDB & ";PWD=" & mvarPassword
Dim UserDataBase As Database
Dim UserDBRecordset As Recordset
Set UserDataBase = OpenDatabase("", False, False, strn)
Set UserDBRecordset = UserDataBase.OpenRecordset("User", dbOpenDynaset)

                      
If aCommand = ulAddUser Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        If RealName = "" Then Call Err.Raise(2, , "Cannot process function without a RealName"): Exit Function
        If flags = "" Then Call Err.Raise(3, , "Cannot process function without Flags"): Exit Function
        If Password = "" Then Call Err.Raise(4, , "Cannot process function without a Password"): Exit Function
        Call AddUser(UserID, RealName, flags, Password)
        UserList = True
ElseIf aCommand = ulEnumerate Then
        For pnCnt = enumCollection.Count To 1 Step -1
            enumCollection.Remove pnCnt
        Next
        While Not UserDBRecordset.EOF
            enumCollection.Add CStr(UserDBRecordset!UserID)
            UserDBRecordset.MoveNext
        Wend
ElseIf aCommand = ulGetFlags Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID = '" & UserID & "'", dbOpenDynaset)
        UserList = UserDBRecordset!access
ElseIf aCommand = ulRemUser Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID= '" & UserID & "'", dbOpenDynaset)
        UserDBRecordset.Delete
        UserList = True
ElseIf aCommand = ulEditUser Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID = '" & UserID & "'", dbOpenDynaset)
        UserDBRecordset.Edit
        If NewID <> "" Then UserDBRecordset!UserID = NewID Else UserDBRecordset!UserID = UserDBRecordset!UserID
        If NewName <> "" Then UserDBRecordset!RealName = NewName Else UserDBRecordset!RealName = UserDBRecordset!RealName
        If NewPass <> "" Then UserDBRecordset!Password = NewPass Else UserDBRecordset!Password = UserDBRecordset!Password
        If NewFlags <> "" Then UserDBRecordset!access = NewFlags Else UserDBRecordset!access = UserDBRecordset!access
        UserDBRecordset.Update
        UserList = True
ElseIf aCommand = ulSetPass Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        If NewPass = "" Then Call Err.Raise(1, , "Cannot process function without a NewPass"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID = '" & UserID & "'", dbOpenDynaset)
        UserDBRecordset.Edit
        UserDBRecordset!Password = NewPass
        UserDBRecordset.Update
        UserList = True
ElseIf aCommand = ulGetPass Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID = '" & UserID & "'", dbOpenDynaset)
        UserList = UserDBRecordset!Password
ElseIf aCommand = ulGetRealName Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID = '" & UserID & "'", dbOpenDynaset)
        UserList = UserDBRecordset!RealName
ElseIf aCommand = ulUserExists Then
        If UserID = "" Then Call Err.Raise(1, , "Cannot process function without a UserID"): Exit Function
        Set UserDBRecordset = UserDataBase.OpenRecordset("SELECT * FROM User WHERE UserID = '" & UserID & "'", dbOpenDynaset)
        If UserDBRecordset.RecordCount = 0 Then UserList = False Else UserList = True
End If
    
Exit Function

err_Init:
        MsgBox "Could not process function. Details follow:" & vbCrLf & Error$, vbCritical, "Error " & Err
        UserList = False
End Function
Function CPVerify(User$, CPID$) As Boolean
cp1$ = CheckPoint(cpGetFlags, CPID)
cpA = ParseCPFlags(ptRequired, cp1$)
cpP = ParseCPFlags(ptProhibit, cp1$)
ulf = UserList(ulGetFlags, User)
CPL = 0 'Length of actual verified checkpoints
CPT = Len(cpA) 'Length of target checpoints
For i = 1 To Len(ulf)
    If InStr(cpP, Mid(ulf, i, 1)) <> 0 Then CPVerify = False: Exit Function
Next i

For i = 1 To Len(ulf)
    If InStr(cpA, Mid(ulf, i, 1)) <> 0 Then CPL = CPL + 1
Next i
If CPL < CPT Then CPVerify = False Else CPVerify = True
End Function



Sub CreateDB(udb, Password)
    Dim WS As DAO.Workspace
    Dim DB As DAO.Database

    Set WS = DAO.DBEngine.Workspaces(0)

    If Dir(udb) = vbNullString Then
        Set DB = WS.CreateDatabase(udb, dbLangGeneral & ";pwd=" & Password)
    Else
        Exit Sub
Exit Sub
    End If

    TD_Access DB
    TD_Checkpoint DB
    TD_User DB

    DB.Close

    Set DB = Nothing
    Set WS = Nothing
End Sub

' Table Access
Private Sub TD_Access(DB As Database)
    Dim TD As DAO.TableDef
    Dim IDX As DAO.Index
    Dim bNewTable As Boolean
    Dim bNewIndex As Boolean

    bNewTable = CreateTable(DB, TD, "Access", 0, "", "", "", "")
    AddField TD, "Flag", 10, 1, False, 2, "", 0, False, "", ""
    AddField TD, "Description", 12, 0, False, 2, "", 1, False, "", ""
    bNewIndex = CreateIndex(TD, IDX, "PrimaryKey", False, True, True, True, False)
    AddIndexField IDX, "Flag", 0
    If bNewIndex Then TD.Indexes.Append IDX
    If bNewTable Then DB.TableDefs.Append TD
    Set IDX = Nothing
    Set TD = Nothing
End Sub

' Table Checkpoint
Private Sub TD_Checkpoint(DB As Database)
    Dim TD As DAO.TableDef
    Dim IDX As DAO.Index
    Dim bNewTable As Boolean
    Dim bNewIndex As Boolean

    bNewTable = CreateTable(DB, TD, "Checkpoint", 0, "", "", "", "")
    AddField TD, "Checkpoint", 10, 50, False, 2, "", 0, False, "", ""
    AddField TD, "Flag", 10, 50, False, 2, "", 1, False, "", ""
    bNewIndex = CreateIndex(TD, IDX, "PrimaryKey", False, True, True, True, False)
    AddIndexField IDX, "Checkpoint", 0
    If bNewIndex Then TD.Indexes.Append IDX
    If bNewTable Then DB.TableDefs.Append TD
    Set IDX = Nothing
    Set TD = Nothing
End Sub

' Table User
Private Sub TD_User(DB As Database)
    Dim TD As DAO.TableDef
    Dim IDX As DAO.Index
    Dim bNewTable As Boolean
    Dim bNewIndex As Boolean

    bNewTable = CreateTable(DB, TD, "User", 0, "", "", "", "")
    AddField TD, "UserID", 10, 255, False, 2, "", 0, False, "", ""
    AddField TD, "RealName", 10, 255, False, 2, "", 1, False, "", ""
    AddField TD, "Access", 10, 255, False, 2, "", 2, False, "", ""
    AddField TD, "Password", 10, 255, False, 2, "", 3, False, "", ""
    bNewIndex = CreateIndex(TD, IDX, "PrimaryKey", False, True, True, True, False)
    AddIndexField IDX, "UserID", 0
    If bNewIndex Then TD.Indexes.Append IDX
    bNewIndex = CreateIndex(TD, IDX, "UserID", False, False, False, False, False)
    AddIndexField IDX, "UserID", 0
    If bNewIndex Then TD.Indexes.Append IDX
    If bNewTable Then DB.TableDefs.Append TD
    Set IDX = Nothing
    Set TD = Nothing
End Sub

Private Function CreateTable(DB As Database, TD As TableDef, TBName As String, _
                             lAttributes As Long, sConnect As String, sSourceTableName As String, _
                             sValidationRule As String, sValidationText As String) As Boolean
    Dim bFound As Boolean
    Dim iInd As Integer

    On Error Resume Next
    bFound = False
    For iInd = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(iInd).Name = TBName Then
            bFound = True
            Exit For
        End If
    Next
    If bFound Then
        Set TD = DB(TBName)
    Else
        Set TD = DB.CreateTableDef(TBName)
        TD.Connect = sConnect
        TD.SourceTableName = sSourceTableName
        TD.Attributes = lAttributes
    End If
    TD.ValidationRule = sValidationRule
    TD.ValidationText = sValidationText
    CreateTable = Not bFound
    On Error GoTo 0
End Function

Private Sub AddField(TD As TableDef, FLDName As String, _
                     iType As Integer, iSize As Integer, _
                     bAllowZeroLenght As Boolean, lAttributes As Long, _
                     sDefaultValue As String, iOrdinalPosition As Integer, _
                     bRequired As Boolean, sValidationRule As String, _
                     sValidationText As String)
    Dim bFound As Boolean
    Dim iInd As Integer
    Dim FLD As DAO.Field

    On Error Resume Next
    bFound = False
    For iInd = 0 To TD.Fields.Count - 1
        If TD.Fields(iInd).Name = FLDName Then
            bFound = True
            Exit For
        End If
    Next
    If bFound Then
        Set FLD = TD.Fields(FLDName)
    Else
        Set FLD = TD.CreateField(FLDName, iType, iSize)
        FLD.Attributes = lAttributes
    End If
    FLD.AllowZeroLength = bAllowZeroLenght
    FLD.DefaultValue = sDefaultValue
    FLD.OrdinalPosition = iOrdinalPosition
    FLD.Required = bRequired
    FLD.ValidationRule = sValidationRule
    FLD.ValidationText = sValidationText
    If Not bFound Then TD.Fields.Append FLD
    Set FLD = Nothing
    On Error GoTo 0
End Sub

Private Function CreateIndex(TD As TableDef, IDX As Index, IDXName As String, _
                             bClustered As Boolean, bPrimary As Boolean, _
                             bUnique As Boolean, bRequired As Boolean, _
                             bIgnoreNulls As Boolean) As Boolean
    Dim bFound As Boolean
    Dim iInd As Integer

    On Error Resume Next
    bFound = False
    For iInd = 0 To TD.Indexes.Count - 1
        If TD.Indexes(iInd).Name = IDXName Then
            bFound = True
            Exit For
        End If
    Next
    If bFound Then
        Set IDX = TD.Indexes(IDXName)
    Else
        Set IDX = TD.CreateIndex(IDXName)
        IDX.Clustered = bClustered
        IDX.Primary = bPrimary
        IDX.Unique = bUnique
        IDX.Required = bRequired
        IDX.IgnoreNulls = bIgnoreNulls
    End If
    CreateIndex = Not bFound
    On Error GoTo 0
End Function

Private Sub AddIndexField(IDX As Index, FLDName As String, lAttributes As Long)
    Dim bFound As Boolean
    Dim iInd As Integer
    Dim FLD As DAO.Field

    On Error Resume Next
    bFound = False
    For iInd = 0 To IDX.Fields.Count - 1
        If IDX.Fields(iInd).Name = FLDName Then
            bFound = True
            Exit For
        End If
    Next
    If bFound Then
        Set FLD = IDX.Fields(FLDName)
    Else
        Set FLD = IDX.CreateField(FLDName)
        FLD.Attributes = lAttributes
    End If
    If Not bFound Then IDX.Fields.Append FLD
    Set FLD = Nothing
    On Error GoTo 0
End Sub

Private Function TableExists(DB As Database, TBName As String) As Boolean
    Dim iInd As Integer

    On Error Resume Next
    TableExists = False
    For iInd = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(iInd).Name = TBName Then
            TableExists = True
            Exit For
        End If
    Next
    On Error GoTo 0
End Function

Private Function IndexExists(TD As TableDef, IDXName As String) As Boolean
    Dim iInd As Integer

    On Error Resume Next
    IndexExists = False
    For iInd = 0 To TD.Indexes.Count - 1
        If TD.Indexes(iInd).Name = IDXName Then
            IndexExists = True
            Exit For
        End If
    Next
    On Error GoTo 0
End Function


