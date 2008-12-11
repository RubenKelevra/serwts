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

Private mstrMboxOldFileName As String 'filled in renameMBox and holds the old
                                      'filename as long as we haven't saved
                                      'the new to the configfile
                                      'after that we will remove the old file
                                      'and delete the file and clear this var again
Private mstrMboxFileName As String
Private mstrMboxPath As String
Private mblnMboxOpen As Boolean
Private mintFnum As Integer
Public parentbox As Integer
Private mstrMBoxLabel As String

Private mMD5 As xclsMD5
Private mSHA256 As xclsSHA256

Public Function getMBoxName() As String
    'Public getMBoxName As String
    ' - returns the name of the mbox
    getMBoxName = mstrMBoxLabel
End Function

Public Function renameMBox(ByRef newName As String) As Boolean
    'Public renameMBox As Boolean
    'renames the MBox, moves the MBox file 'cause we won't find it again if we changed the name
    'further information look at destription of mstrMboxOldFileName
    ' - [IN] ByRef newName As String: gives us the new name
    ' - returns true if we had copied the file probably
    
End Function


'fixme not done
Private Function openHandle() As Boolean
    'Public openMBox As Boolean
    'Open filehandle to MBox
    ' - returns true if this was ok
    If Not mblnMboxOpen Then
        If FileExists(mstrMboxPath & mstrMboxFileName) Then
            mintFnum = FreeFile
            Open mstrMboxPath & mstrMboxFileName For Random As #mintFnum
            mblnMboxOpen = True
            openHandle = True
        Else
            Call Err.Raise(vbObjectError, , "MBox couldn't be found: '" & mstrMboxPath & mstrMboxFileName & "'")
            Exit Function
        End If
    Else
        openHandle = False
    End If
End Function

Public Function isOpen() As Boolean
    isOpen = mblnMboxOpen
End Function

Public Function closeMbox() As Boolean
    'Public closeMBox As Boolean
    'Close filehandle to MBox
    ' - returns true if this was ok
    If mblnMboxOpen Then
        Close #mintFnum
        mblnMboxOpen = False
        closeMbox = True
    Else
        closeMbox = False
    End If
End Function

Public Function openMbox(ByRef Path As String, ByRef label As String) As Boolean
    'this function should open the mbox with the given laben in the given path
End Function

Public Function createMbox(ByRef Path As String, ByRef Filename As String) As Boolean
    'this function should create an empty mbox on the given position, returning
    'false if there is already an mbox or the file could not be created
End Function

Private Function calcFilename(ByRef label As String) As String
    'Private calcFilename As String
    ' - [IN] ByRef Filename As String: label of the mbox
    ' - returns the "encrypted" filename for our mbox
    calcFilename = mMD5.sum(mSHA256.sum(label))
End Function

'fixme not done
Private Function setPathFile(ByRef Path As String, ByRef Filename As String, CreateNonExisting As Boolean) As Boolean
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

Private Sub Class_Initialize()
    Set mSHA256 = New xclsSHA256
    Set mMD5 = New xclsMD5
End Sub
