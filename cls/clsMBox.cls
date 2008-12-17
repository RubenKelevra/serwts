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
Private mstrMboxFromline As String

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
    
    Dim iFileLen As Integer
    Dim iFileLenOld As Integer
    
    iFileLenOld = FileSystem.FileLen(mstrMboxPath)
    
    mstrMboxOldFileName = mstrMboxFileName
    mstrMboxFildeName = calcFilename(newName)
    FileSystem.FileCopy mstrMboxPath, mstrMboxPath
    
    If FileExists(mstrMboxPath) Then
        iFileLen = FileSystem.FileLen(mstrMboxPath)
        If (Len(iFileLenOld) = Len(iFileLen)) Then
            renameMBox = True
        Else
            renameMBox = False
        End If
    Else
        renameMBox = False
    End If
    
    End Function

Private Function openHandle() As Boolean
    'Private openMBox As Boolean
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

Public Function openMBox(ByRef path As String, ByRef label As String) As Integer
    'this function should open the mbox with the given laben in the given path
    Open path & label For Random As #mintFnum
End Function

Public Function createMbox(ByRef path As String, ByRef Filename As String) As Integer
    'this function should create an empty mbox on the given position, returning
    'false if there is already an mbox or the file could not be created
    mstrMboxPath = path
    mstrMboxFileName = Filename
    
    If openHandle Then
        createMbox = Error.Success
    Else
        createMbox = Error.Unsuccess
    End If
End Function

Private Function calcFilename(ByRef label As String) As String
    'Private calcFilename As String
    ' - [IN] ByRef Filename As String: label of the mbox
    ' - returns the "encrypted" filename for our mbox
    calcFilename = mMD5.sum(mSHA256.sum(label)) & ".box"
End Function

Public Function setPathFile(ByRef path As String, ByRef label As String, CreateNonExisting As Boolean) As Boolean
    mstrMboxFileName = calcFilename(label)
    If CreateNonExisting Then
        If Not FileExists(path & mstrMboxFileName) Then
            'creating folder if it isn't existing
            If Not (CBool(createDir(path))) Then
                Call Err.Raise(vbObjectError, , "MBox directory couldn't be created, check filesystem permissions")
                Exit Function
            End If
            If Not createMbox(path, mstrMboxFileName) Then
                Exit Function
            End If
        End If
    Else
        If Not FileExists(path & mstrMboxFileName) Then
            Call Err.Raise(vbObjectError, , "MBox file does not exist")
            Exit Function
        End If
    End If
    mstrMboxPath = path
    Call openHandle
    
End Function

Private Function splitheader(ByRef sheader As String) As Integer
    'splitheader(sheader) as Integer
    'gibt als Integer einen wert fehler wert oder eine ok als zulässig aus
    'This function split the string by ever " " and save the parts in other strings
    '// sTeilstring(0) = "From " if not = false
    '// sTeilstring(1) = "email-Address" if not = false example : myemail@web.de
    '// sTeilstring(2) = "date(Type Date zb. 10.12.2008)
    'Split the Header in From, Absender, Betreff, Datum
    
    Dim sTeilstring() As String
    Dim sFrom As String
    Dim sAddresse As String
    Dim sDate As String
    
    sTeilstring = Split(sheader)
    sFrom = sTeilstring(0)
    sAddresse = sTeilstring(1)
    sDate = sTeilstring(2)
End Function

Private Function buildseperator(ByRef sheader As String) As Integer
    'look for the Fromline and the next line
    'then take both lines in one string for example : From c@d.com date
    Dim FromLine As String
    Dim DateLine As String
    Dim HeaderLine As String
    
    'FromLine reading
    Pos1 = InStr(1, sheader, "FROM:")
    FromLine = Mid(sheader, Pos1 + 5, vbCrLf)
    'example : root <root@arktur.btr-rs.local>
    
    'DateLine reading
    Pos1 = InStr(1, sheader, "DATE:")
    DateLine = Mid(sheader, Pos1 + 5, vbCrLf)
    'example : wed, 26 Nov 2008 07:58:08 +0100
    
    HeaderLine = "From " & FromLine & DateLine
    

End Function
Private Function getsubject(ByRef sheader) As String
    

End Function
Private Function getdate(ByRef sheader) As String
    'DateLine reading
    Pos1 = InStr(1, sheader, "DATE:")
    DateLine = Mid(sheader, Pos1 + 5, vbCrLf)
    'example : wed, 26 Nov 2008 07:58:08 +0100

End Function
Private Function getAddyFrom() As String
    'FromLine reading
    Pos1 = InStr(1, sheader, "FROM:")
    FromLine = Mid(sheader, Pos1 + 5, vbCrLf)
    'example : root <root@arktur.btr-rs.local>
End Function
Private Function getAddyTo() As String
    
End Function
Public Function getMailInfo(detailed As Boolean) As String()
'1.Subject
'2.AddyFrom
'3.Date
'If detailed = true then
'4.AddyTo
'5.Content Type:


End Function


Private Function saveEmailtoMbox(ByRef HeaderLine As String, Header As String, Body As String, path As String, Filename As String) As Integer
    'funtion saveEmailtoMbox(fromline, Header, Body, path, Filename) as integer
    'this function give an integer as error marker or as ok sign back
    'the task is to save first the fromline than the header and then the body
    Open path & Filename For Input As #mintFnum
        Print #mintFnum, HeaderLine
        Print #mintFnum, Header
        Print #mintFnum, Body
        Print #mintFnum, vbCrLf
        Print #mintFnum, vbCrLf
    Close #mintFnum

End Function

Private Sub Class_Initialize()
    Set mSHA256 = New xclsSHA256
    Set mMD5 = New xclsMD5
End Sub
