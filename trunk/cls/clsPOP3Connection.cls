VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPOP3Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Handles one POP3 Connection for Emailfetching

Private mMD5 As xclsMD5 'used for APOP-command
Private mwskNetworkSocket As Winsock
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

Private Sub Class_Initialize()
    'Private Class_Initialize
    'space to run some functions to initialize some (default)values
    Set mMD5 = New xclsMD5
    mblnConfigured = False
End Sub

Private Sub clearMailData()
    Length = -1
    ID = -1
    Mail = ""
    UIDL = ""
End Sub

Public Function Create(ByRef Pass As String, ByRef User As String, ByRef wskNetworkSocket As Winsock, ByRef strServerAddress As String, intPort As Integer, blnSecureAuthorisation) As Integer
    mstrPassword = Pass
    mstrUser = User
    Set mwskNetworkSocket = wskNetworkSocket
    mstrServerAddress = strServerAddress
    mintPort = intPort
    mblnSecureAuthorisation = blnSecureAuthorisation
    mblnConfigured = True
    Call clearMailData
    Create = Error.Success
End Function

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
End Function

Private Function WskSetup() As Boolean
    On Error GoTo WskSetupError
    With mwskNetworkSocket
        .RemoteHost = mstrServerAddress
        .RemotePort = mintPort
        .Protocol = sckTCPProtocol
    End With
    On Error GoTo 0
    WskSetup = True
    Exit Function
    
WskSetupError:
    On Error GoTo 0
    WskSetup = False
    Exit Function
    
End Function

Private Function WskConnect() As Boolean
    On Error GoTo WskConnectError
    mwskNetworkSocket.Connect
    On Error GoTo 0
    WskConnect = True
    
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

Private Function login() As Integer
    Dim i As Integer
    If Not configured Then
        login = Error.NotConfigured
        Exit Function
    End If
    
    If Not WskSetup Then
        login = Error.WskConf
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
        
        'we going to connect to Server
        Call .Connect
        DoEvents
        i = 0
        While .State = sckConnecting
            If i > SckToutConnecting Then
                login = Error.WskConnectingTout
                Exit Function
            End If
            i = i + 10
            'Call Sleep(10)
        Wend
        
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
        
        i = getIncommingData(mstrTempIncommingData)
        'if we don't get an ok, we abort
        If Not Left$(mstrTempIncommingData, 3) = "+OK" Then
            login = Error.PopCommunicationError
            Exit Function
        ElseIf i = Error.WskCommunicationTout Then
            login = Error.WskCommunicationTout
            Exit Function
        End If
        
        mstrLoginTime = "<" + Split(mstrTempIncommingData, "<")(1)
        
        'check syntax
        
        If FindFirstOf(mstrLoginTime, "<") + 1 = FindLastOf(mstrLoginTime, ">") Then
            'no date given
            mstrLoginTime = ""
        ElseIf FindLastOf(mstrLoginTime, ">") = -1 Then
            'no date given
            mstrLoginTime = ""
        End If
        
        On Error GoTo SendDataError
        Call .SendData("USER " & mstrUser)
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
            Call .SendData("PASS " & mstrPassword)
        Else
            Call .SendData("PASS " & mMD5.sum(mstrLoginTime & mstrPassword & vbNullChar))
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

Public Function getMailheaders(ByRef Output() As String) As Integer
    If Not configured Then
        getMailheaders = Error.NotConfigured
        Exit Function
    End If
    
    getMailheaders = login
End Function

Public Function getMail(Number As Integer) As Integer
    If Not configured Then
        checkSecureAuthorisation = Error.NotConfigured
        Exit Function
    End If
End Function

Public Function getMail(Number As Integer) As Integer
    If Not configured Then
        checkSecureAuthorisation = Error.NotConfigured
        Exit Function
    End If
End Function
