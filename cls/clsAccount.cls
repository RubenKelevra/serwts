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

Private Declare Function createDir _
 Lib "imagehlp.dll" (ByVal lpPath As String) As Long
 
Private Declare Sub CopyMemoryPtr Lib "kernel32" _
  Alias "RtlMoveMemory" ( _
  ByVal Destination As Long, _
  ByVal Source As Long, _
  ByVal Length As Long)


Private Const ConfHeaderLines = 3
Private Const ConfBodyLines = 7

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
Public mAccountSettings As clsAccountSettings

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
        Call Err.Raise(vbObjectError, , "Application directory couldn't be created, check filesystem rights")
    End If
    
    Set mSHA256 = New xclsSHA256
    Set mBlowfish = New xclsBlowfish
    Set mHuffman = New xclsHuffman
    Set mMD5 = New xclsMD5
    Set mAccountSettings = New clsAccountSettings
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
        If Len(Data) = 0 Or Data Is Nothing Then 'there is no acc file
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
        Erase Data
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
                .mstrServer & vbLf & _
                CStr(.mintPort) & vbLf & _
                .mstrUsername & vbLf & _
                .mstrPassword & vbLf & _
                .mstrMboxPath & vbLf & _
                CStr(.mdtmLastFetch) & vbLf
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
    For i = LBound(Data) To UBound(Data) - LBound(Data)
        If foundEnd Then
            ReDim Preserve AccSettings(CInt((elements - ConfHeaderLines) / ConfBodyLines))
            Exit For
        End If
        If Data(i) = vbLf Then
            If element = 0 Then
                'this is the SHA256 of user, ignoring
            ElseIf element = 1 Then
                'this is the (default) Maildir
                mMailDir = Mid(Data, lastSeperator, i - lastSeperator)
                If mMailDir = "" Then
                    mMailDir = mPath & "Mail" & BackSlash
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
                    currentAccount = CInt((elements - ConfHeaderLines) / ConfBodyLines)
                    If currentAccount > reservedSpace Then
                        reservedSpace = reservedSpace + 10
                        ReDim Preserve AccSettings(reservedSpace)
                    End If
                    Call AccSettings(currentAccount) _
                            .saveData((elements - ConfHeaderLines) Mod ConfBodyLines, partData)
                End If
            End If
                
            element = element + 1
            lastSeperator = i
        End If
    Next i
End Sub

Public Sub SaveConfig(ByRef AccSettings() As clsAccountSettings)
    'Private SaveConfig
    'savely writes all configdata encrypted to disk, should be called
    'on every change of configurations
    ' - [IN]ByRef AccSettings() As clsAccountSettings: existing accounts
    '   informations
    Dim Data As String
    Dim fnum As Integer
    fnum = FreeFile
    
    'plaintext to huffman compressed, blowfish encrypted, base64 encoded data
    Data = mBlowfish.EncryptString(mHuffman.EncodeString(AccountWriter(AccSettings)), mPass, True)
    
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
    Erase AccountSettings: LastLoginTime: mUser: mPass: mUserChkSum
End Sub

Public Sub logout()
    'Public logout
    'we going to save to disk and destroy in memory existing data for a save logout
    If Not mLoggedin Then
        Exit Sub
    End If
    Call SaveConfig(mAccountSettings)
    Call clearData
    
    mLoggedin = False
End Sub

Public Function getAccountnames() As String()
    'Public getAccountnames As String()
    'Returns a stringarray which holds all accountnames
    ' - all accountnames
    Dim i As Integer
    For i = LBound(mAccSettings) To UBound(mAccSettings) - LBound(mAccSettings)
        Call getAccountnames.Join(mAccountSettings(i).mstrAccountName)
    Next i
End Function

Public Function deleteAccountSettings(ByRef index As Long) As Boolean
    'Public deleteAccountSetting As Boolean
    'delete one specified accountsettings from array
    ' - [IN] ByRef index as Long: position which has to be deleted
    ' - true if AccountSettings has successfully deleted
    Dim lngPtr As Long 'pointer var
    
    If (Not mLoggedin) Or (index < LBound(mAccountSettings)) _
            Or (UBound(mAccountSettings) - LBound(mAccountSettings) < index) Then
        Exit Function
    End If
    If index < UBound(mAccountSettings) - LBound(mAccountSettings) Then 'we don't have to delete the last element
        'now we have to copy the data after deleted element over it
        lngPtr = StrPtr(mAccountSettings(index))
        
        Call CopyMemoryPtr(lngPtr, VarPtr(lngPtr + 1), VarPtr(UBound(mAccountSettings) - LBound(mAccountSettings)) - VarPtr(index))
    End If
    ReDim Preserve mAccountSettings(UBound(mAccountSettings) - LBound(mAccountSettings) - 1)
    
    'we changed our configuration, we have to save it to disk
    Call SaveConfig(mAccountSettings)
    
    deleteAccountSettings = True
End Function

