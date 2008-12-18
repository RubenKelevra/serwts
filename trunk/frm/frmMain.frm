VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Pop-Client"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin MSWinsockLib.Winsock wskMain 
      Left            =   480
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3000
      Width           =   4575
   End
   Begin MSComctlLib.ListView lsvTestView 
      Height          =   2535
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMD5 As xclsMD5 'used for APOP-command
Private mstrServerAddress As String
Private mintPort As Integer
Private mblnSecureAuthorisation As Boolean
Private mblnConfigured As Boolean
Private mstrPassword As String
Private mstrUser As String
Private mstrLoginTime As String 'for APOP command, if not set there's no APOP capability
Private mstrTempIncommingData As String 'saves temporary the incomming data
Public Mail As String
Public UID As String
Public Length As Long
Public ID As Long
Private currentNetworkStatus As Integer
Private currentPOP3Task As Integer

Private Sub Command1_Click()
    
    If wskMain.State = sckClosed Then
        wskMain.RemoteHost = "pop.northpolyptica.de"
        wskMain.RemotePort = 110
        mstrUser = "testing@northpolyptica.de"
        mstrPassword = "pop"
        
        currentNetworkStatus = POP3Stat.awaitingFirstOK
        currentPOP3Task = POP3TaskCode.getEmailHeaders
        Call wskMain.Connect
    Else
    End If
    
End Sub

Private Sub Form_Load()
    With lsvTestView
        .ColumnHeaders.Add , , "From"
        .ColumnHeaders.Add , , "Subject"
        .ColumnHeaders.Add , , "Date"
        
        .View = lvwReport
        .ListItems.Add , , "zzA"
        .ListItems.Item(1).ListSubItems.Add , , "ffA"
        .ListItems.Item(1).ListSubItems.Add , , "ddA"
        
        .ListItems.Add , , "ttB"
        .ListItems.Item(2).ListSubItems.Add , , "rrB"
        .ListItems.Item(2).ListSubItems.Add , , "ssB"
    End With
    
    'init of vars
    Set mMD5 = New xclsMD5
    mblnConfigured = False
    currentNetworkStatus = POP3Stat.closed
End Sub

Private Sub lsvTestView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDescending(2) As Boolean

    If ColumnHeader.index = 1 Then 'there it starts with 1!
        'MsgBox "first one"
    ElseIf ColumnHeader.index = 2 Then
        'MsgBox "second one"
    End If
    
    With lsvTestView
        .SortKey = ColumnHeader.index - 1 'there it starts with 0!
        .SortOrder = Switch(blnDescending(ColumnHeader.index - 1), lvwDescending, True, lvwAscending)
        .Sorted = True
        blnDescending(ColumnHeader.index - 1) = Not blnDescending(ColumnHeader.index - 1)
    End With
    
    
End Sub

Private Sub clearMailData()
    Length = -1
    ID = -1
    Mail = ""
    UID = ""
End Sub

Private Function configured() As Boolean
    configured = mblnConfigured
End Function

Public Function secureAuthorisation(Optional value As Object) As Integer
    'Public secureAuthorisation As Boolean
    'use this to set or get the secure authorisation status
    ' - [IN] Optional Value As Object: if set it will override the state
    ' - will return the current setting, furthermore it will return the Value state if set
    If Not configured Then
        secureAuthorisation = Error.NotConfigured
        Exit Function
    End If
    If IsMissing(value) Then
        secureAuthorisation = CInt(mblnSecureAuthorisation)
    ElseIf getVarType(value) = 3 Then
        mblnSecureAuthorisation = CBool(value)
        secureAuthorisation = CInt(mblnSecureAuthorisation)
    End If
    
    
    Call WskConnect
    
End Function

Private Function WskConnect() As Boolean
    On Error GoTo WskConnectError
    wskMain.Connect
    On Error GoTo 0
    WskConnect = True
    Exit Function
    
WskConnectError:
    On Error GoTo 0
    WskConnect = False
    Exit Function
    
End Function

Public Function checkSecureAuthorisation() As Integer
    'Public secureAuthorisation As Integer
    'returns the possibility of a SecureAuthorisation to the POP3 Server
    ' - returns Error.Success if possible
    If Not configured Then
        checkSecureAuthorisation = Error.NotConfigured
        Exit Function
    End If
    
    If Not WskSetup Then
        checkSecureAuthorisation = Error.WskConf
    End If
    
    If Not WskConnect Then
        checkSecureAuthorisation = Error.WskConnect
    End If
    
End Function

Private Function CalcAPOP_MD5(ByRef Arrived_Greeting_Data As String, ByRef Pass As String) As String
        'if we don't get an ok, we abort
        If Not Left$(Arrived_Greeting_Data, 3) = "+OK" Then
            Exit Function
        End If
        
        'we wanna get the timestamp
        CalcAPOP_MD5 = "<" + Split(Arrived_Greeting_Data, "<")(1)
        CalcAPOP_MD5 = Mid(CalcAPOP_MD5, 1, Len(CalcAPOP_MD5) - 2)
        
        'checking syntax
        If InStr(CalcAPOP_MD5, "<") + 1 = InStr(CalcAPOP_MD5, ">") Then
            'no date given
            CalcAPOP_MD5 = ""
        ElseIf Not Right(CalcAPOP_MD5, 1) = ">" Then
            'no date given
            CalcAPOP_MD5 = ""
        Else
            CalcAPOP_MD5 = LCase$(mMD5.sum(CalcAPOP_MD5 & mstrPassword & vbNullChar))
        End If
End Function

Private Function login() As Integer
    Dim i As Integer
    If Not configured Then
        login = Error.NotConfigured
        Exit Function
    End If
      
    With mwskNetworkSocket
        'if connection not closed, we'll close it hardly
        If Not .State = sckClosed Then
            Call .Close
            i = 0
            'checking if connection is closed
            While .State = sckClosing
                If i > SckToutClosing Then
                    login = Error.WskClosingTout
                    Exit Function
                End If
                i = i + 10
                'Call Sleep(10)
            Wend
        End If
        
        'now the connection should be in closed state
        If Not .State = sckClosed Then
            login = Error.WskClosingError
            Exit Function
        End If
                    
        If Not WskSetup Then
            login = Error.WskConf
            Exit Function
        End If

        'we going to connect to Server
        Call .Connect
        i = 0
        While .State = sckConnecting
            DoEvents
            If i > SckToutConnecting Then
                login = Error.WskConnectingTout
                Exit Function
            End If
            i = i + 10
            'Call Sleep(10)
        Wend
        
        i = getIncommingData(mstrTempIncommingData)
        
        'now the connection should be in connected state
        If Not .State = sckConnected Then
            Select Case .State
                Case sckConnectionRefused
                    login = Error.WskConnectionRefused
                Case sckConnectionReset
                    login = Error.WskConnectionReset
                Case sckHostNotFound
                    login = Error.WskHostNotFound
                Case sckNetworkUnreachable
                    login = Error.WskNetworkUnreachable
                Case sckNetworkSubsystemFailed
                    login = Error.WskNetworkSubsystemFailed
                Case Else
                    login = Error.WskConnectingFailed
            End Select
            Exit Function
        End If
        
        'send login, save logintime (for APOP) if present
        

        
        On Error GoTo SendDataError
        Call .SendData("USER " & mstrUser & vbCrLf)
        i = getIncommingData(mstrTempIncommingData)
        If Not Left$(mstrTempIncommingData, 3) = "+OK" Then
            login = Error.PopUserNotFound
            Exit Function
        ElseIf i = Error.WskCommunicationTout Then
            login = Error.WskCommunicationTout
            Exit Function
        End If
        
        If mstrLoginTime = "" And mblnSecureAuthorisation Then
            login = Error.PopAPOPNotSupported
            Exit Function
        ElseIf mstrLoginTime = "" Then
            Call .SendData("PASS " & mstrPassword & vbCrLf)
        Else
            Call .SendData("PASS " & mMD5.sum(mstrLoginTime & mstrPassword & vbNullChar) & vbCrLf)
        End If
        
        i = getIncommingData(mstrTempIncommingData)
        If Not Left$(mstrTempIncommingData, 3) = "+OK" Then
            login = Error.PopPasswordNotCorrect
            Exit Function
        ElseIf i = Error.WskCommunicationTout Then
            login = Error.WskCommunicationTout
            Exit Function
        End If
        
        On Error GoTo 0
    End With
    
    login = Error.Success
    Exit Function
    
SendDataError:
            On Error GoTo 0
            login = Error.WskSendDataError
            Exit Function
End Function

Private Function getIncommingData(ByRef IncommingData As String) As Integer
    Dim i As Integer
    
    i = 0
    While mwskNetworkSocket.BytesReceived = 0
        If i > SckToutData Then
            getIncommingData = Error.WskCommunicationTout
            Exit Function
        End If
        i = i + 10
        'call Sleep (10)
    Wend
    getIncommingData = Error.Success
    Call mwskNetworkSocket.GetData(IncommingData)
End Function

Public Function getMailheaders(ByRef output() As String) As Integer
    If Not configured Then
        getMailheaders = Error.NotConfigured
        Exit Function
    End If
    
    getMailheaders = login
End Function

Public Function getMail(Number As Integer) As Integer
    If Not configured Then
        getMail = Error.NotConfigured
        Exit Function
    End If
End Function

Public Function getMailUID(Number As Integer) As Integer
    If Not configured Then
        getMailUID = Error.NotConfigured
        Exit Function
    End If
End Function

Private Function isOK(ByRef str As String) As Boolean
    If Left$(str, 3) = "+OK" Then
        isOK = True
    Else
        isOK = False
    End If
End Function

Private Sub errorhandler(code As Integer)

End Sub

Private Function getNumbers(ByRef str As String) As Currency()
    Dim splitstr() As String
    Dim i As Integer
    Dim i2 As Integer
    Dim output() As Currency
    Dim savedValues As Integer
    
    splitstr = Split(Mid(str, 1, Len(str) - 2), " ")
    For i = LBound(splitstr) + 1 To UBound(splitstr) - LBound(splitstr)
        If splitstr(i) = "" Then
            'ignore
        Else
            i2 = getVarType(splitstr(i))
            If i2 = 98 Or i2 = 1 Then 'is a 0 or a positive number
                ReDim Preserve output(savedValues + 1)
                output(savedValues + 1) = CCur(splitstr(i))
            End If
            'else ignore not a number
        End If
    Next i
    getNumbers = output
End Function

Private Sub wskMain_DataArrival(ByVal bytesTotal As Long)
    Dim ArrivedData As String
    Dim iArg() As Currency 'currency is used to handle very big numbers
    Dim tempstr() As String
    
    With wskMain
    
        Call .GetData(ArrivedData)
        
        Select Case currentNetworkStatus
            Case POP3Stat.closed 'if we do not expect data, we exit
                'fixme: for debug only
                MsgBox "Unexpected data arrival: '" & ArrivedData & "'"
                Exit Sub
            Case POP3Stat.awaitingFirstOK 'we wait for first ok after connect
                If currentPOP3Task = POP3TaskCode.NoTask Then
                    'but no task has been set, so we quit again
                    currentNetworkStatus = POP3Stat.awaitingQuitOK
                    Call .SendData("QUIT" & vbCrLf)
                Else 'in all other cases we want to login, so we send user or apop
                    If False Then 'If secureAuthorisation Then 'fixme
                        If Not CalcAPOP_MD5(ArrivedData, mstrPassword) = "" Then 'timestamp ok, sending APOP
                            currentNetworkStatus = POP3Stat.awaitingPassOK
                            Call .SendData("APOP " & mstrUser & " " & CalcAPOP_MD5(ArrivedData, mstrPassword) & vbCrLf)
                        Else 'no timestamp has been send by server, so we have to quit
                            Call errorhandler(Error.PopAPOPNotSupported)
                            currentNetworkStatus = POP3Stat.awaitingQuitOK
                            currentPOP3Task = POP3TaskCode.NoTask
                            Call .SendData("QUIT" & vbCrLf)
                        End If
                    Else 'sending plain password authorisation
                        currentNetworkStatus = POP3Stat.awaitingUserOK
                        Call .SendData("USER " & mstrUser & vbCrLf)
                    End If
                End If
            Case POP3Stat.awaitingUserOK 'we wait for an ok after sending user information
                If currentPOP3Task = POP3TaskCode.NoTask Then
                    currentNetworkStatus = POP3Stat.awaitingQuitOK
                    Call .SendData("QUIT" & vbCrLf)
                Else
                    If isOK(ArrivedData) Then
                        currentNetworkStatus = POP3Stat.awaitingPassOK
                        Call .SendData("PASS " & mstrPassword & vbCrLf)
                    Else
                        Call errorhandler(Error.PopPasswordNotCorrect)
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    End If
                End If
            Case POP3Stat.awaitingPassOK 'we wait for an ok after sending password or apop command
                If isOK(ArrivedData) Then
                    Select Case currentPOP3Task
                        Case POP3TaskCode.NoTask
                            currentNetworkStatus = POP3Stat.awaitingQuitOK
                            Call .SendData("QUIT" & vbCrLf)
                        Case POP3TaskCode.checkAPOPCapability
                            Call errorhandler(Error.PopAPOPisSupported)
                            currentNetworkStatus = POP3Stat.awaitingQuitOK
                            Call .SendData("QUIT" & vbCrLf)
                        Case Else 'check if there are mails in the maildir
                            currentNetworkStatus = POP3Stat.awaitingStatOK
                            Call .SendData("STAT" & vbCrLf)
                    End Select
                Else
                    Call errorhandler(Error.PopPasswordNotCorrect)
                    currentPOP3Task = POP3TaskCode.NoTask
                    currentNetworkStatus = POP3Stat.closed
                    wskMain.Close
                End If
            Case POP3Stat.awaitingStatOK
                If currentPOP3Task = POP3TaskCode.NoTask Then
                    currentNetworkStatus = POP3Stat.awaitingQuitOK
                    Call .SendData("QUIT" & vbCrLf)
                ElseIf currentPOP3Task = POP3TaskCode.checkAPOPCapability Then
                    currentNetworkStatus = POP3Stat.awaitingQuitOK
                    Call .SendData("QUIT" & vbCrLf)
                ElseIf isOK(ArrivedData) Then
                    iArg = getNumbers(ArrivedData)
                    If Not UBound(iArg) - LBound(iArg) = 1 Then GoTo staterr 'not (only) two numbers were sended by the server
                    
                    If Not iArg(0) = 0 Then 'there are emails which we going to fetch
                        If Not (currentPOP3Task = POP3TaskCode.checkAPOPCapability Or currentPOP3Task = POP3TaskCode.NoTask) Then
                            currentPOP3Task = POP3Stat.awaitingListOK
                            Call .SendData("LIST" & vbCrLf)
                        Else
                            currentPOP3Task = POP3TaskCode.NoTask
                            currentNetworkStatus = POP3Stat.awaitingQuitOK
                            Call .SendData("QUIT" & vbCrLf)
                        End If
                    Else 'there is nothing to do
                        Call errorhandler(Error.PopNoEmailsInAccount)
                        currentPOP3Task = POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    End If
                Else '-ERR'
staterr:
                    'fixme: add code for this case
                    currentPOP3Task = POP3TaskCode.NoTask
                    currentNetworkStatus = POP3Stat.awaitingQuitOK
                    .SendData ("QUIT" & vbCrLf)
                    'stat doesn't seems to be supported
                    'trying to handle it with list
                End If
            Case POP3Stat.awaitingListOK
                'fixme if is wrong
                If 3 <= currentPOP3Task And currentPOP3Task <= 9 Then 'is with UID
                    currentNetworkStatus = POP3Stat.awaitingUidlOK
                    Call .SendData("UIDL" & vbCrLf)
                ElseIf (10 <= currentPOP3Task And currentPOP3Task <= 16) Or currentNetworkStatus = 2 Then 'without UID or without restrictions
                    currentpop3 = POP3Stat.awaitingListOK
                    .SendData ("LIST" & vbCrLf)
                End If
                        
                        
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingRetrOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingDeleOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingNoopOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingRsetOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingQuitOK
                If isOK(ArrivedData) Then
                    Call wskMain.Close
                Else
                    Call errorhandler(Error.PopErrWhileQuitting)
                    Call wskMain.Close
                End If
            Case POP3Stat.awaitingTopOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingUidlOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case POP3Stat.awaitingApopOK
                Select Case currentPOP3Task
                    Case POP3TaskCode.NoTask
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.checkAPOPCapability
                        currentNetworkStatus = POP3Stat.awaitingQuitOK
                        Call .SendData("QUIT" & vbCrLf)
                    Case POP3TaskCode.getEmailHeaders
                    
                    Case POP3TaskCode.getNewEmailHeaders_Wuid
                    Case POP3TaskCode.getNewEmailHeaders_WOuid
                    Case POP3TaskCode.getOneEmail_Wuid_Wdelete
                    Case POP3TaskCode.getOneEmail_WOuid_Wdelete
                    Case POP3TaskCode.getOneEmail_Wuid_WOdelete
                    Case POP3TaskCode.getOneEmail_WOuid_WOdelete
                    Case POP3TaskCode.deleteEmail_Wuid
                    Case POP3TaskCode.deleteEmail_WOuid
                    Case POP3TaskCode.getAllEmails_Wuid_deleteALL
                    Case POP3TaskCode.getAllEmails_WOuid_deleteALL
                    Case POP3TaskCode.getAllEmails_Wuid_NOdelete
                    Case POP3TaskCode.getAllEmails_WOuid_NOdelete
                    Case POP3TaskCode.getAllEmails_Wuid_deleteFETCHED
                    Case POP3TaskCode.getAllEMails_WOuid_deleteFETCHED
                End Select
            Case Default
                Exit Sub
        End Select
    End With
End Sub

