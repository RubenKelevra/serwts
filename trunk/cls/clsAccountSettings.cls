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

Public Sub saveData(ByRef index As Integer, ByRef data As String)
    'Public saveData
    'sets various values
    ' - [IN] ByRef index As Integer: the index of information which is given
    ' - [IN] ByRef data As String: the information which should be stored
    Select Case index
        Case 0 'Account name
            mstrAccountName = data
        Case 1 'Server
            mstrSever = data
        Case 2 'Port
            If getVarType(data) = 1 Then
                If CInt(data) <= 65536 Then
                    mintPort = CInt(data)
                Else
                    mintPort = 0
                End If
            Else
                mintPort = 0
            End If
        Case 3 'Username
            mstrUsername = data
        Case 4 'Password
            mstrPassword = data
        Case 5 'MBox Path
            mstrMboxPath = data
        Case 6 'Last fetch time
            If getVarType(data) = 4 Or getVarType(data) = 5 Or getVarType(data) = 98 Then
                mdtmLastFetch = CDate(data)
            Else
                mdtmLastFetch = Nothing
            End If
        Case Else
    End Select
End Sub

