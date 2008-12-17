Attribute VB_Name = "modStrFunctions"
Option Explicit

Public Function FindFirstOf(ByRef str As String, ByRef chars As String) As Integer
    Dim s As String
    s = Split(str, chars)
    FindFirstOf = Len(s(0)) + 1
End Function

Public Function FindLastOf(ByRef str As String, ByRef chars As String) As Integer
    Dim s As String
    Dim i As String
    s = Split(str, chars)
    If Not IsArray(s) Then
        FindLastOf = -1
        Exit Function
    End If
    For i = 0 To UBound(s) - LBound(s)
        FindLastOf = FindLastOf + Len(s(i))
    Next i
    FindLastOf = FindLastOf + UBound(s) - LBound(s) - 1
End Function
