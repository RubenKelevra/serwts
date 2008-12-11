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

Dim mMD5 As xclsMD5 'used for APOP-command
Dim mwskNetworkSocket As Winsock
Dim mstrServerAddress As String
Dim mintPort As Integer
Dim mblnSecureAuthorisation As Boolean
Dim mblnConfigured As Boolean

Private Sub Class_Initialize()
    'Private Class_Initialize
    'space to run some functions to initialize some (default)values
    Set mMD5 = New xclsMD5
    mblnConfigured = False
End Sub

Public Function create(ByRef wskNetworkSocket As Winsock, ByRef strServerAddress As String, intPort As Integer, blnSecureAuthorisation) As Integer
    Set mwskNetworkSocket = wskNetworkSocket
    mstrServerAddress = strServerAddress
    mintPort = intPort
    mblnSecureAuthorisation = blnSecureAuthorisation
    mblnConfigured = True
    create = Error.Success
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
        checkSecureAuthorisation = Error.NotConfigured
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
        Call .Close
        .BytesReceived = 0
        .RemoteHost = mstrServerAddress
        .RemotePort = mintPort
        .Protocol = sckTCPProtocol
    End With
    On Error GoTo 0
    WskSetup = True
    
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
        checkSecureAuthorisation = Error.ErrorWskConf
    End If
    If Not WskConnect Then
        checkSecureAuthorisation = Error.WskConnect
    End If
    
    
        
End Function

Public Function getMailheaders(ByRef Output() As String) As Integer
    If Not configured Then
        checkSecureAuthorisation = Error.NotConfigured
        Exit Function
    End If
End Function

Public Function getMail(Number As Integer, ByRef Output As String) As Integer
    If Not configured Then
        checkSecureAuthorisation = Error.NotConfigured
        Exit Function
    End If
End Function
