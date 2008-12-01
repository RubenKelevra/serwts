VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccountSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mstrAccountName As String
Public mstrServer As String
Public mintPort As Integer
Public mstrUsername As String
Public mstrPassword As String
Public mstrMboxPath As String
Public mdtmLastFetch As Date
Public mblnForceSecuredAuthorisation As Boolean

Public Sub saveData(ByRef index As Integer, ByRef Data As String)
    'Public saveData
    'sets various values
    ' - [IN] ByRef index As Integer: the index of information which is given
    ' - [IN] ByRef data As String: the information which should be stored
    Select Case index
        Case 0 'Account name
            mstrAccountName = Data
        Case 1 'Server
            mstrSever = Data
        Case 2 'Port
            If getVarType(Data) = 1 Then
                If CInt(Data) <= 65536 Then
                    mintPort = CInt(Data)
                Else
                    mintPort = 0
                End If
            Else
                mintPort = 0
            End If
        Case 3 'Username
            mstrUsername = Data
        Case 4 'Password
            mstrPassword = Data
        Case 5 'MBox Path
            mstrMboxPath = Data
        Case 6 'Last fetch time
            If getVarType(Data) = 4 Or getVarType(Data) = 5 Or getVarType(Data) = 98 Then
                mdtmLastFetch = CDate(Data)
            Else
                mdtmLastFetch = Nothing
            End If
        Case 7 'force secured authorisation
            mblnForceSecuredAuthorisation = CBool(Data)
        Case Else
    End Select
End Sub

