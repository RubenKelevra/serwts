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
Public mstrPOP3Server As String
Public mintPort As Integer
Public mstrUsername As String
Public mstrPassword As String
Public mstrMboxPath As String
Public mdtmLastFetch As Date
Public mblnForceSecuredAuthorisation As Boolean
'FIXME: this need get/set functions
Private maMBoxes() As clsMBox
Private i As Integer

Public Sub saveData(ByRef index As Integer, ByRef Data As String)
    'Public saveData
    'sets various values
    ' - [IN] ByRef index As Integer: the index of information which is given
    ' - [IN] ByRef data As String: the information which should be stored
    Dim str() As String
    Select Case index
        Case 0 'Account name
            mstrAccountName = Data
        Case 1 'Server
            mstrPOP3Server = Data
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
                mdtmLastFetch = 0
            End If
        Case 7 'force secured authorisation
            mblnForceSecuredAuthorisation = CBool(Data)
        Case 8 'MBox filenames
            str = Split(Data, ";")
            For i = LBound(str) To UBound(str) - LBound(str) Step 3
                Call AddSubbox(CInt(str(i)), str(i + 1), CInt(str(i + 2)))
            Next i
        Case Else
    End Select
End Sub

Private Function AddSubbox(Parent As Integer, ByRef Selfname As String, selfno As Integer)
    
End Function

Public Function insertSubbox(Parent As Integer, ByRef Selfname As String) As Integer
    'Public insertSubbox As Integer
    'is used to add a subfolder mbox to the given folder
    ' - [IN] Parent As Integer: gives the ID from the parentfolder
    ' - [IN] Selfname As String: labels the new folder
    ' - returns the new ID of the folder, if not possible it will return 0
    
End Function
