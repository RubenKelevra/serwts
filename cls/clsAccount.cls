VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ConfHeaderLines = 3
Private Const ConfBodyLines = 8

Private mPath As String     'Our configuration directory which holds all
                            'encrypted configuration files
Private mMailDir As String  'Directory where we save our mails
Private mUser As String     'The given username - if it's set we had logged in
Private mPass As String     'The password which can be used to open the
                            'config-file of the user (not plaintext -
                            'SHA-256 checksummed with username as SHA-256 as salt
Private mFile As String     'The encrypted file's filename which is the MD5 sum
                            'of the SHA-256 sum of mUser
Private mBakFile As String  'The backup file's filename which is the same filename
                            'as mFile but with .bak extention
Private mLoggedin As Boolean 'true if we have opened the configation file successfully
Public LastLoginTime As Date
Private mUserChkSum As String
Private mAccountSettings() As clsAccountSettings

Private mSHA256 As xclsSHA256
Private mBlowfish As xclsBlowfish
Private mHuffman As xclsHuffman
Private mMD5 As xclsMD5

Private Sub Class_Initialize()
    'Private Class_Initialize
    'space to run some functions to initialize some (default)values
    
    'fill path to our application data folder
    mPath = GetSpecialFolder(CSIDL_APPDATA, True, True) + BackSlash + "Serwts Mailclient" + BackSlash
    
    'creating folder if it isn't existing
    If Not (CBool(createDir(mPath))) Then
        Call Err.Raise(vbObjectError, , "Application directory couldn't be created, check filesystem permissions")
    End If
    
    Set mSHA256 = New xclsSHA256
    Set mBlowfish = New xclsBlowfish
    Set mHuffman = New xclsHuffman
    Set mMD5 = New xclsMD5
End Sub

Private Sub SetFilenames(ByRef User As String)
    'Private setFilename
    'calcs the local var for the filenames
    ' - [IN] ByRef User As String: The username which is used to login
    mFile = mMD5.sum(mSHA256.sum(User))
    mBakFile = mFile & ".bak"
    mFile = mFile & ".acc"
End Sub

Private Sub SetPassword(ByRef UserSHA256 As String, ByRef Pass As String)
    'Private setPassword
    'calcs the password out of username and given password which is actual used
    'while decrypting and encrypting via blowfish
    ' - [IN] ByRef User As String: The login name of the user
    ' - [IN] ByRef Pass As String: The given password for this user
    mPass = mSHA256.sum(Pass & UserSHA256)
End Sub
Private Sub SetUserSHA256(ByRef User As String)
    'Private Sub SetUserSHA256
    'calcs the SHA256 sum of the given user and setting the var
    ' - [IN] ByRef User As String: The login name of the user
    mUserChkSum = mSHA256.sum(User)
End Sub

Public Function login(ByRef User As String, ByRef Pass As String) As Boolean
    'Public login As Boolean
    'init all values needed to ask for the accountsettings of this useraccount
    'for this we need to decrypt our configuration file for this user
    ' - [IN] ByRef User As String: The username which is used to login
    ' - [IN] ByRef Pass As String: The given password for this user
    ' - returns true if login was possible
    Dim Data As String
    Dim fnum As Integer
    fnum = FreeFile
    
    'fill the filename var
    Call SetFilenames(User)
    
    On Error Resume Next
        'read the encrypted file from disk
        Open mPath & mFile For Input As #fnum
            Data = Input$(LOF(fnum), fnum)
        Close #fnum
    On Error GoTo fileerror
        If Len(Data) = 0 Or Data = "" Then 'there is no acc file
            fnum = FreeFile
            'try to read backup (bak) file
            Open mPath & mBakFile For Input As #fnum
                Data = Input$(LOF(fnum), fnum)
            Close #fnum
        End If
    On Error GoTo 0
    
    'fill SHA256 sum var
    Call SetUserSHA256(User)
    'fill password
    Call SetPassword(mUserChkSum, Pass)
    
    'base64 encoded, blowfish encrypted, huffman compressed data to plaintext
    Data = mHuffman.DecodeString(mBlowfish.DecryptString(Data, mPass, True))
    
    If Left(Data, Len(mUserChkSum)) = mUserChkSum Then 'is the file successfully decrypted
        mUser = User
        Call AccountReader(Data, mAccountSettings)
        'this var isn't needed anymore
        Data = ""
        mLoggedin = True
    Else
        mLoggedin = False
        login = False
        Exit Function
    End If
    
fileerror:
    On Error GoTo 0
    mLoggedin = False
    login = False
    Exit Function
    
End Function

Private Function AccountWriter(ByRef AccSettings() As clsAccountSettings) As String
    'Private AccountWriter As String
    'build the filestring which has to be crypted out of AccSettings
    ' - [IN] ByRef AccSettings() As clsAccountSettings: Source for the file
    ' - Final filestring
    Dim i As Integer
    AccountWriter = mUserChkSum & vbLf & mMailDir & vbLf & CStr(LastLoginTime) & vbLf
    For i = LBound(AccSettings) To UBound(AccSettings) - LBound(AccSettings)
        With AccSettings(i)
            AccountWriter = AccountWriter & _
                .mstrAccountName & vbLf & _
                .mstrPOP3Server & vbLf & _
                CStr(.mintPort) & vbLf & _
                .mstrUsername & vbLf & _
                .mstrPassword & vbLf & _
                .mstrMboxPath & vbLf & _
                CStr(.mdtmLastFetch) & vbLf & _
                CStr(.mblnForceSecuredAuthorisation) & vbLf
        End With
    Next i
    AccountWriter = AccountWriter & "END" & vbEOF
End Function

Private Sub AccountReader(ByRef Data As String, ByRef AccSettings() As clsAccountSettings)
    'Private AccountReader
    'Reads the account data from a decrypted string to AccSettings
    ' - [IN] ByRef Data As String: Decrypted data as string
    ' - [IN] ByRef AccSettings() As clsAccountSettings: Target location
    Dim i As Integer
    Dim lastSeperator As Integer
    Dim element As Integer
    Dim partData As String
    Dim foundEnd As Boolean
    Dim reservedSpace As Integer
    Dim currentAccount As Integer
    'reserve space for first 10 accounts (and destroy existing data)
    reservedSpace = 10
    ReDim AccSettings(reservedSpace)
    'running thru all characters
    For i = 0 To Len(Data) - 1
        If foundEnd Then
            ReDim Preserve AccSettings(CInt((element - ConfHeaderLines) / ConfBodyLines))
            Exit For
        End If
        If Mid(Data, i, 1) = vbLf Then
            If element = 0 Then
                'this is the SHA256 of user, ignoring
            ElseIf element = 1 Then
                'this is the (default) Maildir
                mMailDir = Mid(Data, lastSeperator, i - lastSeperator)
                If mMailDir = "" Then
                    mMailDir = getMailDir("", "")
                End If
            ElseIf element = 2 Then
                'this is the last login time
                LastLoginTime = CDate(Mid(Data, lastSeperator, i - lastSeperator))
            Else
                partData = Mid(Data, lastSeperator, i - lastSeperator)
                If partData = "END" Then
                    foundEnd = True
                ElseIf partData = "-END" Then
                    partData = "END"
                Else
                    currentAccount = CInt((element - ConfHeaderLines) / ConfBodyLines)
                    If currentAccount > reservedSpace Then
                        reservedSpace = reservedSpace + 10
                        ReDim Preserve AccSettings(reservedSpace)
                    End If
                    Call AccSettings(currentAccount) _
                            .saveData((element - ConfHeaderLines) Mod ConfBodyLines, partData)
                End If
            End If
                
            element = element + 1
            lastSeperator = i
        End If
    Next i
End Sub

Private Sub SaveConfig()
    'Private SaveConfig
    'savely writes all configdata encrypted to disk, should be called
    'on every change of configurations
    ' - [IN]ByRef AccSettings() As clsAccountSettings: existing accounts
    '   informations
    Dim Data As String
    Dim fnum As Integer
    fnum = FreeFile
    
    'plaintext to huffman compressed, blowfish encrypted, base64 encoded data
    Data = mBlowfish.EncryptString(mHuffman.EncodeString(AccountWriter(mAccountSettings)), mPass, True)
    
    'check if there is already a backupfile, if so we delete it
    If FileExists(mPath & mBakFile) Then
        Call FileSystem.Kill(mPath & mBakFile)
    End If
    
    'if acc file exists
    If FileExists(mPath & mFile) Then
        'copy acc file as backup file
        Call FileSystem.FileCopy(mPath & mFile, mPath & mBakFile)
        'delete acc file
        Call FileSystem.Kill(mPath & mFile)
    End If
    
    'write data to acc file
    Open mPath & mFile For Output As #fnum
        Put fnum, , Data
    Close #fnum
End Sub

Private Sub clearData()
    'Private clearData
    'erase all data for logout
    Erase mAccountSettings
    LastLoginTime = 0
    mUser = ""
    mPass = ""
    mUserChkSum = ""
End Sub

Public Sub logout()
    'Public logout
    'we going to save to disk and destroy in memory existing data for a save logout
    If Not mLoggedin Then
        Exit Sub
    End If
    Call SaveConfig
    Call clearData
    
    mLoggedin = False
End Sub

Public Function getAccountnames() As String()
    'Public getAccountnames As String()
    'Returns a stringarray which holds all accountnames
    ' - all accountnames
    Dim i As Integer
    Dim str() As String
    ReDim str(UBound(mAccountSettings) - LBound(mAccountSettings))
    For i = LBound(mAccountSettings) To UBound(mAccountSettings) - LBound(mAccountSettings)
         str(i) = mAccountSettings(i).mstrAccountName
    Next i
    getAccountnames = str
End Function

Private Function checkASIndex(ByRef index As Long) As Boolean
    If (Not mLoggedin) Or (index < LBound(mAccountSettings)) _
            Or (UBound(mAccountSettings) - LBound(mAccountSettings) < index) Then
        Exit Function
    End If
    checkASIndex = True
End Function

Public Function deleteAccountSettings(ByRef index As Long) As Boolean
    'Public deleteAccountSetting As Boolean
    'delete one specified accountsettings from array
    ' - [IN] ByRef index as Long: position which has to be deleted
    ' - true if AccountSettings has successfully deleted
    Dim lngPtr As Long 'pointer var
    
    If Not checkASIndex(index) Then
        Exit Function
    End If
    
    If index < UBound(mAccountSettings) - LBound(mAccountSettings) Then 'we don't have to delete the last element
        'now we have to copy the data after deleted element over it
        lngPtr = StrPtr(mAccountSettings(index))
        
        Call CopyMemory(lngPtr, VarPtr(lngPtr + 1), VarPtr(UBound(mAccountSettings) - LBound(mAccountSettings)) - VarPtr(index))
    End If
    ReDim Preserve mAccountSettings(UBound(mAccountSettings) - LBound(mAccountSettings) - 1)
    
    'we changed our configuration, we have to save it to disk
    Call SaveConfig
    
    deleteAccountSettings = True
End Function


Public Sub addAccountSettings(ByRef Accountname As String, ByRef POP3Server As String, ByRef Port As Integer, ByRef Username As String, ByRef Password As String, ForceSecuredAuthorisation As Boolean, Optional ByRef MboxPath As String = "")
    'Public addAccountSettings
    'we add a new mailaccount
    ' - [IN] ByRef AccountName As String: The displayed name of the account
    ' - [IN] ByRef POP3Server As String: The POP3 server address
    ' - [IN] ByRef Port As Integer: Port of the server 0 means default
    ' - [IN] ByRef Username As String: The username which is used to login
    ' - [IN] ByRef Password As String: The password which is used to login
    ' - [IN] Optional ByRef MboxPath As String = "":
    'allocate space for new account
    ReDim Preserve mAccountSettings(UBound(mAccountSettings) - LBound(mAccountSettings) + 1)
    
    With mAccountSettings(UBound(mAccountSettings) - LBound(mAccountSettings))
        .mstrAccountName = Accountname
        .mstrPOP3Server = POP3Server
        .mintPort = Port
        .mstrUsername = Username
        .mstrPassword = Password
        .mstrMboxPath = MboxPath
        .mdtmLastFetch = 0
        .mblnForceSecuredAuthorisation = ForceSecuredAuthorisation
    End With
    
    Call SaveConfig
End Sub

Private Function mailFetched(ByRef index As Long) As Boolean
    'Private mailFetched As Boolean
    'renew date/time of the given mail account
    ' - [IN] ByRef index As Long: index of the mail account
    ' - returns true if everything is ok
    If Not checkASIndex(index) Then
        Exit Function
    End If

    mAccountSettings(index).mdtmLastFetch = DateTime.Now
    
    Call SaveConfig
    
    mailFetched = True
End Function

Public Function changeMailboxName(ByRef index As Long, ByRef newName As String) As Boolean
    'Public changeMailboxName As Boolean
    'change the name of the given mail account
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] ByRef newName As String: the new name if the mail account
    ' - returns true if everything is ok
    Dim i As Integer
    If Not checkASIndex(index) Then
        Exit Function
    End If
    
    If mAccountSettings(index).mstrAccountName = newName Then 'we have nothing to do
        changeMailboxName = True
        Exit Function
    End If
    
    'name of the mailaccount is the primary key, duplicates aren't allowed
    'we have to check this
    For i = LBound(mAccountSettings) To UBound(mAccountSettings) - LBound(mAccountSettings)
        If (Not i = index) And _
                Trim$(LCase$(mAccountSettings(i).mstrAccountName)) = _
                    Trim$(LCase$(newName)) Then
            Call Err.Raise(vbObjectError, , "The given name not unique, use a different name for your account")
            Exit Function
        End If
    Next i
    
    mAccountSettings(index).mstrAccountName = newName
    
    Call SaveConfig
    
    changeMailboxName = True
End Function

Public Function changeMailUser(ByRef index As Long, ByRef newUsername As String) As Boolean
    'Public changeMailUser As Boolean
    'change the name of the user which is used to login to the given mail account
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] ByRef newUsername As String: the new username
    ' - returns true if everything is ok
    If Not checkASIndex(index) Then
        Exit Function
    End If
    If Not mAccountSettings(index).mstrUsername = newUsername Then
        mAccountSettings(index).mstrUsername = newUsername
        
        Call SaveConfig
    End If
    changeMailUser = True
End Function

Public Function changeMailPassword(ByRef index As Long, ByRef newPassword As String) As Boolean
    'Public changeMailPassword As Boolean
    'change the password which is used to login to the given mail account
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] ByRef newPassword As String: The new password
    ' - returns true if everything is ok
    If Not checkASIndex(index) Then
        Exit Function
    End If
    If Not mAccountSettings(index).mstrPassword = newPassword Then
        mAccountSettings(index).mstrPassword = newPassword
        
        Call SaveConfig
    End If
    changeMailPassword = True
End Function

Public Function changePOP3Server(ByRef index As Long, ByRef newPOP3Server As String) As Boolean
    'Public changePOP3Server As Boolean
    'change the POP3Server which is used fetch mails of the given mailaccount
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] ByRef newPOP3Server As String: the new POP3 server address
    ' - returns true if everything is ok
    If Not checkASIndex(index) Then
        Exit Function
    End If
    If Not mAccountSettings(index).mstrPOP3Server = newPOP3Server Then
        mAccountSettings(index).mstrPOP3Server = newPOP3Server
        
        Call SaveConfig
    End If
    changePOP3Server = True
End Function

Public Function changeForceSecuredAuthorisation(ByRef index As Long, newVal As Boolean) As Boolean
    'Public changeForceSecuredAuthorisation As Boolean
    'change the state of forcing secured authorisation of the given mailaccount
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] ByRef newVal As Boolean: The new state of this switch
    ' - returns true if everything is ok
    If Not checkASIndex(index) Then
        Exit Function
    End If
    If Not mAccountSettings(index).mblnForceSecuredAuthorisation = newVal Then
        mAccountSettings(index).mblnForceSecuredAuthorisation = newVal
        
        Call SaveConfig
    End If
    changeForceSecuredAuthorisation = True
End Function

Public Function changePOP3ServerPort(ByRef index As Long, newPort As Integer) As Boolean
    'Public changePOP3ServerPort As Boolean
    'change the port of the POP3 server of the given mailaccount
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] newPort As Integer: the new port, 0 means default
    ' - returns true if everything is ok
    If Not checkASIndex(index) Then
        Exit Function
    End If
    If Not mAccountSettings(index).mintPort = newPort Then
        mAccountSettings(index).mintPort = newPort
    
        Call SaveConfig
    End If
    changePOP3ServerPort = True
End Function

Private Function getMailDir(ByRef MboxPath As String, ByRef Accountname As String) As String
    If MboxPath = "" Then
        getMailDir = mPath & "Mail" & BackSlash
    Else
        getMailDir = MboxPath
        If Not Right(MboxPath, 1) = BackSlash Then
            getMailDir = getMailDir & BackSlash
        End If
    End If
    If Not Accountname = "" Then
        getMailDir = getMailDir & mMD5.sum(mSHA256.sum(Accountname)) & BackSlash
    End If
End Function

Private Function closeMBox(ByRef index As Long)
    
End Function

Private Function openMbox(ByRef index As Long)
    
End Function

Private Function moveMbox(ByRef index As Long, ByRef newPath As String, ByRef newAccountname As String)
    Dim mboxfile As Variant
    Dim sourceDirectory As String
    Dim destinationDirectory As String
    
    With mAccountSettings(index)
        If .mstrAccountName = newAccountname Then
            If .mstrMboxPath = newPath Then 'there is nothing to do
                moveMbox = True
                Exit Function
            End If
        End If
    
        'get old and new path
    
        sourceDirectory = getMailDir(.mstrMboxPath, .mstrAccountName)
        destinationDirectory = getMailDir(newPath, newAccountname)
    End With
    
    'check if directory is equal
    If sourceDirectory = destinationDirectory Then 'there is nothing to do
        moveMbox = True
        Exit Function
    End If
    
    Call closeMBox(index)
    For Each mboxfile In getMboxFiles(index)
        If Not FileExists(sourceDirectory & mboxfile) Then
            Call Err.Raise(vbObjectError, , "Error mboxfile " & mboxfile & " couldn't be found")
            'fixme: clean up
            Exit Function
        End If
        'file seems to be there, we try to copy
        On Error GoTo filecopyerror
        Call FileSystem.FileCopy(sourceDirectory & mboxfile, destinationDirectory & mboxfile)
        On Error GoTo 0
    Next mboxfile
    
    For Each mboxfile In getMboxFiles(index)
        If Not FileExists(destinationDirectory & mboxfile) Then
            Call Err.Raise(vbObjectError, , "Error on copy, destination file does not exist anymore: " & mboxfile & "")
            'fixme: clean up
            Exit Function
        End If
        'destination file seems to be ok
        On Error GoTo sourcefiledeleteerror
        Call FileSystem.Kill(sourceDirectory & mboxfile)
        On Error GoTo 0
    Next mboxfile
    
filecopyerror:
    Call Err.Raise(vbObjectError, , "File couldn't be copy to destination directory or couldn't be read" & vbCrLf & "Source: '" & sourceDirectory & mboxfile & "'" & vbCrLf & "Destination: '" & destinationDirectory & mboxfile & "'")
    Exit Function
sourcefiledeleteerror:
    Call Err.Raise(vbObjectError, , "Sourcefile couldn't be deleted" & mboxfile)
    Exit Function
    
End Function

Private Function getMboxFiles(ByRef index As Long) As String()
    
End Function

Public Function changeMboxPath(ByRef index As Long, ByRef newPath As String) As Boolean
    'Public changeMboxPath As Boolean
    'change the path of the mbox of the given mailaccount
    ' - [IN] ByRef index As Long: index of the mail account
    ' - [IN] ByRef newPath As String: the new path of the mbox, "" means use default directory
    ' - returns true if everything is ok

    
    If Not checkASIndex(index) Then
        Exit Function
    End If
    
    
        
End Function
