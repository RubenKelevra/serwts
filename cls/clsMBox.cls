VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrMboxFileName As String
Private mstrMboxPath As String
Private mblnMboxOpen As Boolean
Private mintFnum As Integer

Public Function openMBox() As Boolean
    'Public openMBox As Boolean
    'Open filehandle to MBox
    ' - returns true if this was ok
    If Not mblnMboxOpen Then
        If FileExists(mstrMboxPath & mstrMboxFileName) Then
            mintFnum = FreeFile
            Open mstrMboxPath & mstrMboxFileName For Random As #mintFnum
            mblnMboxOpen = True
            openMBox = True
        Else
            Call Err.Raise(vbObjectError, , "MBox couldn't be found: '" & mstrMboxPath & mstrMboxFileName & "'")
            Exit Function
        End If
    Else
        openMBox = False
    End If
End Function

Public Function isOpen() As Boolean
    isOpen = mblnMboxOpen
End Function

Public Function closeMBox() As Boolean
    'Public closeMBox As Boolean
    'Close filehandle to MBox
    ' - returns true if this was ok
    If mblnMboxOpen Then
        Close #mintFnum
        mblnMboxOpen = False
        closeMBox = True
    Else
        closeMBox = False
    End If
End Function

Private Function createMbox(ByRef Path As String, ByRef Filename As String) As Boolean
    'this function should create an empty mbox on the given position, returning
    'false if there is already an mbox or the file could not be created
End Function
Public Function setPathFile(ByRef Path As String, ByRef Filename As String, CreateNonExisting As Boolean) As Boolean
    If CreateNonExisting Then
        If Not FileExists(Path & Filename) Then
            'creating folder if it isn't existing
            If Not (CBool(createDir(Path))) Then
                Call Err.Raise(vbObjectError, , "MBox directory couldn't be created, check filesystem permissions")
                Exit Function
            End If
            If Not createMbox(Path, Filename) Then
                Exit Function
            End If
        End If
    Else
        If Not FileExists(Path & Filename) Then
            Call Err.Raise(vbObjectError, , "MBox file does not exist")
            Exit Function
        End If
    End If
    mstrMboxPath = Path
    mstrMboxFileName = Filename
End Function
