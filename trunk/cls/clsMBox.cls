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
'ready
Public Function getMBoxName() As String
    'Public getMBoxName As String
    ' - returns the name of the mbox
    getMBoxName = mstrMBoxLabel
End Function
'ready
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
'ready
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
'ready
Public Function isOpen() As Boolean
    isOpen = mblnMboxOpen
End Function
'ready
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
'ready
Public Function openMBox(ByRef path As String, ByRef label As String) As Integer
    'this function should open the mbox with the given laben in the given path
    Open path & label For Random As #mintFnum
End Function
'ready
Public Function createMbox(ByRef path As String, ByRef filename As String) As Integer
    'this function should create an empty mbox on the given position, returning
    'false if there is already an mbox or the file could not be created
    mstrMboxPath = path
    mstrMboxFileName = filename
    
    If openHandle Then
        createMbox = Error.Success
    Else
        createMbox = Error.Unsuccess
    End If
End Function
'ready
Private Function calcFilename(ByRef label As String) As String
    'Private calcFilename As String
    ' - [IN] ByRef Filename As String: label of the mbox
    ' - returns the "encrypted" filename for our mbox
    calcFilename = mMD5.sum(mSHA256.sum(label)) & ".box"
End Function
'ready
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
'ready
Private Function splitheader(ByRef sheaderline As String) As Integer
    'splitheader(sheader) as Integer
    'gibt als Integer einen Fehler Wert oder eine Ok als zulässig aus
    'This function split the string by ever " " and save the parts in other strings
    '// sTeilstring(0) = "From " if not = false
    '// sTeilstring(1) = "email-Address" if not = false example : myemail@web.de
    '// sTeilstring(2) = "date(Type Date zb. 10.12.2008)
    'Split the Header in From, Absender, Betreff, Datum
    
    Dim sTeilstring() As String
    Dim sFrom As String
    Dim sAddresse As String
    Dim sDate As String
    
    sTeilstring = Split(sheaderline, " ")
    
    sFrom = sTeilstring(0)
    sAddresse = sTeilstring(1)
    sDate = sTeilstring(2)
End Function
'ready
Private Function buildseperator(ByRef sheader As String) As Integer
    ' take both lines in one string for example : From c@d.com date
    Dim FromLine As String
    Dim DateLine As String
    Dim HeaderLine As String
    
    FromLine = getAddyFrom
    DateLine = getDate
    HeaderLine = "From " & FromLine & DateLine
    

End Function
'ready
Private Function getSubject(ByRef sheader) As String
    'subjectLine reading
    pos1 = InStr(1, sheader, "SUBJECT:")
    getSubject = Mid(sheader, pos1 + 8, vbCrLf)
    'example : Das ist ein Beispiel Betreff
    
End Function
'ready
Private Function getDate(ByRef sheader) As String
    'DateLine reading
    pos1 = InStr(1, sheader, "DATE:")
    getDate = Mid(sheader, pos1 + 5, vbCrLf)
    'example : wed, 26 Nov 2008 07:58:08 +0100

End Function
'ready
Private Function getAddyFrom(ByRef sheader) As String
    'FromLine reading
    pos1 = InStr(1, sheader, "FROM:")
    getAddyFrom = Mid(sheader, pos1 + 5, vbCrLf)
    'example : root <root@arktur.btr-rs.local>
End Function
'ready
Private Function getAddyTo(ByRef sheader) As String
    'ToLine reading
    pos1 = InStr(1, sheader, "To:")
    getAddyTo = Mid(sheader, pos1 + 3, vbcrfl)
End Function
'ready
Private Function getUID(sheader) As String
    'reading the line where the UID is
    pos1 = InStr(1, sheader, "UID:")
    getUID = Mid(sheader, pos1 + 4, vbcrfl)
End Function
'ready
Private Function getIndex(sheader) As String
    'reading the line where the Index is
    pos1 = InStr(1, sheader, "x-Index:")
    getIndex = Mid(sheader, pos1 + 8, vbcrfl)
End Function

'not ready
Public Function getMailInfo(detailed As Boolean) As String()
        'save this in string arry by ever index
        '1.Subject
        '2.AddyFrom
        '3.Date
        'If detailed = true then
        '4.AddyTo
        '5.Content Type:
        If detailed = True Then
            
        Else
            
        End If
End Function

'not ready
Private Function addheader(ByRef iIndex As Integer) As String
    'the task of this function is to call only the header
    

End Function

'not ready
Private Function replaceaddmail(ByRef iIndex As Integer) As String
    'private function replaceaddmail as string
    'the task of this function is to replace the the old header in the mbox and save the header & email there
    'look after the index and save the new header & Email
    Dim path As String
    Dim filename As String
    Dim sSearchIndex As String
    
    sSearchIndex = sSearchIndex & iIndex
    Open path & filename For Input As #mintFnum
        pos1 = InStr(1, filename, sSearchIndex)
        
    Close #mintFnum
    
End Function
'ready
Public Function addEmail(sEmail) As String()
'Public function()as string()
'the coming Email should split in header an body

addEmail = Split(sEmail, "")
'addEmail(0) = header
'addEmail(1) = Body



End Function
'ready
Private Function saveEmailtoMbox(ByRef sEmail As String) As Integer
    'funtion saveEmailtoMbox(fromline, Header, Body, path, Filename) as integer
    'this function give an integer as error marker or as ok sign back
    'the task is to save first the fromline than the header and then the body
    
    Dim iFileLen As Integer
    Dim HeaderLine As String
    Dim Header As String
    Dim Body As String
    Dim path As String
    Dim filename As String
    Static mintMboxIndex As Integer
    
    
    path = mstrMboxPath
    filename = mstrMboxFileName
    hederline = buildseperator(sEmail)
    Header = addEmail(sEmail)(0)
    Body = addEmail(sEmail)(1)
    mintMboxIndex = mintMboxIndex + 1
    iFileLen = FileSystem.FileLen(mstrMboxPath)
    
    Open path & filename For Input As #mintFnum
        Print #mintFnum, HeaderLine
        Print #mintFnum, vbCrLf
        Print #mintFnum, Header
        Print #mintFnum, vbCrLf
        Print #mintFnum, Body
        Print #mintFnum, vbCrLf
        Print #mintFnum, "x-Index: "
        Print #mintFnum, mintMboxIndex
        Print #mintFnum, vbCrLf
        Print #mintFnum, "x-Size: "
        Print #mintFnum, iFileLen
        Print #mintFnum, vbCrLf
        Print #mintFnum, vbCrLf
    Close #mintFnum

End Function

Private Sub Class_Initialize()
    Set mSHA256 = New xclsSHA256
    Set mMD5 = New xclsMD5
End Sub
