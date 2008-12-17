Attribute VB_Name = "xmodFileOperations"
Option Explicit
 
'Sleep API call
Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
 
'DLL call to create a path recursive, and check on low level if path exists
Public Declare Function createDir _
 Lib "imagehlp.dll" (ByVal lpPath As String) As Long
 
'DLL call to copy a part of memory which is specified to another location via pointer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

'Author is "RobDog888" from www.vbforums.com

Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 255
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
 
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
 
Public Function FileExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
    
End Function

